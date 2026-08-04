"""LLM router — Anthropic, Gemini, Groq, or OpenCode Zen via env LLM_PROVIDER."""
from __future__ import annotations

import json
import os
import re
from typing import Any

import httpx
import tenacity

LLM_PROVIDER = os.getenv("LLM_PROVIDER", "anthropic").strip().lower()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "YOUR_ANTHROPIC_API_KEY_HERE")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY_HERE")
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "YOUR_GROQ_API_KEY_HERE")
OPENCODE_API_KEY = os.getenv("OPENCODE_API_KEY", "")
OPENCODE_MODEL = os.getenv("OPENCODE_MODEL", "deepseek-v4-flash-free")
OPENCODE_FALLBACK_MODELS = [
    m.strip()
    for m in os.getenv(
        "OPENCODE_FALLBACK_MODELS",
        "deepseek-v4-flash-free,big-pickle",
    ).split(",")
    if m.strip()
]
LLM_FALLBACK_TO_OPENCODE = os.getenv("LLM_FALLBACK_TO_OPENCODE", "true").strip().lower() in (
    "true",
    "1",
    "yes",
)


def _opencode_skip_response_format(model: str) -> bool:
    """Free Zen models often reject json_object / json_schema."""
    m = (model or "").lower()
    return any(x in m for x in ("deepseek", "nemotron", "free", "big-pickle"))

DEFAULT_COMPILER_MODEL = "claude-sonnet-4-6"
DEFAULT_GEMINI_MODEL = "gemini-2.0-flash"
DEFAULT_GROQ_MODEL = "llama-3.3-70b-versatile"
OPENCODE_ZEN_CHAT_URL = "https://opencode.ai/zen/v1/chat/completions"

_TIMEOUT = 120.0


def _key_ok(key: str) -> bool:
    return bool(key) and not key.startswith("YOUR_")


def _opencode_key() -> str:
    if _key_ok(OPENCODE_API_KEY):
        return OPENCODE_API_KEY
    if LLM_PROVIDER == "opencode" and _key_ok(ANTHROPIC_API_KEY):
        return ANTHROPIC_API_KEY
    return ""


def _anthropic_quota_exceeded(status: int, detail: str) -> bool:
    d = (detail or "").lower()
    if status not in (400, 402, 403, 429):
        return False
    return any(
        x in d
        for x in (
            "usage limit",
            "rate limit",
            "quota",
            "billing",
            "credit balance",
            "insufficient",
        )
    )


def _opencode_fallback_model(requested: str | None) -> str:
    """When falling back from Anthropic, ignore claude-* model ids."""
    req = (requested or "").strip()
    if req and "claude" not in req.lower():
        return req
    return OPENCODE_MODEL


def is_configured() -> bool:
    if LLM_PROVIDER == "gemini":
        return _key_ok(GEMINI_API_KEY)
    if LLM_PROVIDER == "groq":
        return _key_ok(GROQ_API_KEY)
    if LLM_PROVIDER == "opencode":
        return bool(_opencode_key())
    if LLM_PROVIDER == "openai":
        return _key_ok(os.getenv("OPENAI_API_KEY", ""))
    if LLM_PROVIDER in ("mock", "none"):
        return False
    return _key_ok(ANTHROPIC_API_KEY)


def _repair_json_text(text: str) -> str:
    """Best-effort fixes for common LLM JSON mistakes (no extra dependencies)."""
    s = (text or "").strip()
    if not s:
        return s
    s = s.replace("\u201c", '"').replace("\u201d", '"').replace("\u2018", "'").replace("\u2019", "'")
    # Trailing commas before } or ]
    s = re.sub(r",\s*([}\]])", r"\1", s)
    # Missing comma between objects in arrays: "}\n  {" -> "},\n  {"
    s = re.sub(r"\}\s*\{", "},{", s)
    # Missing comma between array elements that are strings: '"\n  "' -> '",\n  "'
    s = re.sub(r'"\s*\n\s*"', '",\n"', s)
    return s


def _extract_json(text: str) -> Any:
    text = (text or "").strip()
    if not text:
        raise ValueError("empty LLM response")
    candidates: list[str] = [text]
    fence = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if fence:
        candidates.insert(0, fence.group(1).strip())
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        candidates.append(text[start : end + 1])

    last_err: json.JSONDecodeError | None = None
    for raw in candidates:
        for attempt_text in (raw, _repair_json_text(raw)):
            if not attempt_text:
                continue
            try:
                return json.loads(attempt_text)
            except json.JSONDecodeError as e:
                last_err = e
                continue
    if last_err:
        raise last_err
    raise ValueError("no JSON object in LLM response")


def salvage_compiler_payload(raw: str) -> dict[str, Any]:
    """Extract scope / prompts from broken JSON using brace matching."""
    out: dict[str, Any] = {}
    text = (raw or "").strip()
    if not text:
        return out

    def _extract_object(key: str) -> dict | list | None:
        pat = re.compile(rf'"{re.escape(key)}"\s*:\s*(\{{|\[)', re.I)
        m = pat.search(text)
        if not m:
            return None
        open_ch = m.group(1)
        close_ch = "}" if open_ch == "{" else "]"
        start = m.start(1)
        depth = 0
        for i in range(start, len(text)):
            c = text[i]
            if c == open_ch:
                depth += 1
            elif c == close_ch:
                depth -= 1
                if depth == 0:
                    chunk = text[start : i + 1]
                    for candidate in (chunk, _repair_json_text(chunk)):
                        try:
                            return json.loads(candidate)
                        except json.JSONDecodeError:
                            continue
                    return None
        return None

    scope = _extract_object("scope")
    if isinstance(scope, dict):
        out["scope"] = scope
    for key in ("funnel_prompts", "discovery_prompts", "prompts"):
        val = _extract_object(key)
        if isinstance(val, list) and val:
            out[key] = val
            break
    return out


def _stub_complete(_system: str, _user: str) -> str:
    return json.dumps({"status": "placeholder"})


_RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}


def _is_retryable_llm_error(exc: BaseException) -> bool:
    if isinstance(exc, httpx.HTTPStatusError):
        return exc.response.status_code in _RETRYABLE_STATUS_CODES
    return isinstance(exc, (httpx.ConnectError, httpx.NetworkError, OSError))


@tenacity.retry(
    retry=tenacity.retry_if_exception(_is_retryable_llm_error),
    wait=tenacity.wait_exponential(multiplier=1, min=2, max=60),
    stop=tenacity.stop_after_attempt(5),
    reraise=True,
)
def _post_json(url: str, headers: dict, body: dict) -> dict:
    from vendor_intel.clients.http_proxy import httpx_client

    with httpx_client(timeout=_TIMEOUT, skip_pool=True) as client:
        r = client.post(url, headers=headers, json=body)
        r.raise_for_status()
        return r.json()


_ZERO_USAGE: dict[str, int] = {"input_tokens": 0, "output_tokens": 0}


def _usage_anthropic(data: dict[str, Any]) -> dict[str, int]:
    u = data.get("usage") or {}
    return {
        "input_tokens": int(u.get("input_tokens") or 0),
        "output_tokens": int(u.get("output_tokens") or 0),
    }


def _usage_openai_like(data: dict[str, Any]) -> dict[str, int]:
    """openai / groq / opencode all speak the OpenAI chat-completions usage shape."""
    u = data.get("usage") or {}
    return {
        "input_tokens": int(u.get("prompt_tokens") or 0),
        "output_tokens": int(u.get("completion_tokens") or 0),
    }


def _usage_gemini(data: dict[str, Any]) -> dict[str, int]:
    u = data.get("usageMetadata") or {}
    return {
        "input_tokens": int(u.get("promptTokenCount") or 0),
        "output_tokens": int(u.get("candidatesTokenCount") or 0),
    }


def _anthropic_complete(system: str, user: str, model: str, max_tokens: int) -> tuple[str, dict[str, int]]:
    data = _post_json(
        "https://api.anthropic.com/v1/messages",
        {
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
        {
            "model": model,
            "max_tokens": max_tokens,
            "system": system,
            "messages": [{"role": "user", "content": user}],
        },
    )
    text = "".join(p.get("text", "") for p in data.get("content") or [] if p.get("type") == "text")
    return text, _usage_anthropic(data)


def _opencode_message_text(message: dict[str, Any]) -> str:
    """deepseek-v4-flash-free often fills reasoning_content and leaves content empty."""
    content = str(message.get("content") or "").strip()
    if content:
        return content
    reasoning = str(message.get("reasoning_content") or "").strip()
    if not reasoning:
        return ""
    # Some responses embed the final JSON only at the end of the reasoning trace.
    if "{" in reasoning:
        start = reasoning.rfind("{")
        end = reasoning.rfind("}")
        if end > start:
            return reasoning[start : end + 1].strip()
    return ""


def _opencode_effective_max_tokens(max_tokens: int) -> int:
    """Reasoning models burn completion tokens on thinking before the visible answer."""
    return max(max_tokens, 512)


def _gemini_complete(system: str, user: str, model: str) -> tuple[str, dict[str, int]]:
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    data = _post_json(
        url,
        {"x-goog-api-key": GEMINI_API_KEY, "content-type": "application/json"},
        {
            "systemInstruction": {"parts": [{"text": system}]},
            "contents": [{"role": "user", "parts": [{"text": user}]}],
            "generationConfig": {"temperature": 0.2},
        },
    )
    usage = _usage_gemini(data)
    candidates = data.get("candidates") or []
    if not candidates:
        return "", usage
    text = "".join(p.get("text", "") for p in candidates[0].get("content", {}).get("parts") or [])
    return text, usage


def _opencode_complete(system: str, user: str, model: str, max_tokens: int) -> tuple[str, dict[str, int]]:
    """OpenCode Zen — OpenAI-compatible chat API (free models e.g. deepseek-v4-flash-free)."""
    api_key = _opencode_key()
    headers = {"Authorization": f"Bearer {api_key}", "content-type": "application/json"}
    token_budget = _opencode_effective_max_tokens(max_tokens)

    def _call(body_tokens: int, extra: dict[str, Any]) -> tuple[str, str, dict[str, int]]:
        base_body: dict[str, Any] = {
            "model": model,
            "max_tokens": body_tokens,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "temperature": 0.1,
            **extra,
        }
        data = _post_json(OPENCODE_ZEN_CHAT_URL, headers, base_body)
        usage = _usage_openai_like(data)
        choices = data.get("choices") or []
        if not choices:
            return "", "", usage
        choice = choices[0]
        msg = choice.get("message") or {}
        return _opencode_message_text(msg), str(choice.get("finish_reason") or ""), usage

    if _opencode_skip_response_format(model):
        extras: list[dict[str, Any]] = [{}, {"response_format": {"type": "json_object"}}]
    else:
        extras = [{"response_format": {"type": "json_object"}}, {}]
    for extra in extras:
        try:
            text, finish, usage = _call(token_budget, extra)
            if not text and finish == "length" and token_budget < 16384:
                text, _, usage = _call(min(token_budget * 2, 16384), extra)
            if text:
                return text, usage
        except httpx.HTTPStatusError as e:
            if extra and e.response.status_code in (400, 422):
                continue
            if e.response.status_code >= 500:
                return (
                    json.dumps(
                        {
                            "error": f"http_{e.response.status_code}: {e.response.text[:200]}",
                            "status": "llm_failed",
                        }
                    ),
                    dict(_ZERO_USAGE),
                )
            raise
    return (
        json.dumps(
            {
                "error": "empty_response: model used token budget on reasoning; retry or raise max_tokens",
                "status": "llm_failed",
            }
        ),
        dict(_ZERO_USAGE),
    )


def _groq_complete(system: str, user: str, model: str, max_tokens: int) -> tuple[str, dict[str, int]]:
    data = _post_json(
        "https://api.groq.com/openai/v1/chat/completions",
        {"Authorization": f"Bearer {GROQ_API_KEY}", "content-type": "application/json"},
        {
            "model": model,
            "max_tokens": max_tokens,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "temperature": 0.1,
            "response_format": {"type": "json_object"},
        },
    )
    usage = _usage_openai_like(data)
    choices = data.get("choices") or []
    text = str(choices[0].get("message", {}).get("content", "")) if choices else ""
    return text, usage


def _openai_complete(system: str, user: str, model: str, max_tokens: int) -> tuple[str, dict[str, int]]:
    api_key = os.getenv("OPENAI_API_KEY") or ""
    base_url = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").strip()
    headers = {"Authorization": f"Bearer {api_key}", "content-type": "application/json"}
    body = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
        "temperature": 0.2,
    }
    if "gpt" in model.lower():
        body["response_format"] = {"type": "json_object"}
    url = f"{base_url.rstrip('/')}/chat/completions"
    data = _post_json(url, headers, body)
    usage = _usage_openai_like(data)
    choices = data.get("choices") or []
    text = str(choices[0].get("message", {}).get("content", "")) if choices else ""
    return text, usage


def _resolve_dispatch_model(provider: str, model: str | None) -> str:
    if provider == "gemini":
        return model or DEFAULT_GEMINI_MODEL
    if provider == "groq":
        return model or DEFAULT_GROQ_MODEL
    if provider == "opencode":
        return model or OPENCODE_MODEL
    if provider == "openai":
        return model or "gpt-4o-mini"
    return model or DEFAULT_COMPILER_MODEL


def _dispatch(
    system: str,
    user: str,
    *,
    model: str | None,
    max_tokens: int,
) -> tuple[str, dict[str, int], str, str]:
    """Core provider dispatch — returns (text, usage, provider_used, model_used).

    provider_used/model_used reflect what ACTUALLY served the request, which
    can differ from LLM_PROVIDER/model when the Anthropic-quota-exceeded ->
    OpenCode fallback kicks in (important for accurate cost attribution).
    """
    if not is_configured():
        return _stub_complete(system, user), dict(_ZERO_USAGE), "mock", ""
    try:
        if LLM_PROVIDER == "gemini":
            resolved = _resolve_dispatch_model("gemini", model)
            text, usage = _gemini_complete(system, user, resolved)
            return text, usage, "gemini", resolved
        if LLM_PROVIDER == "groq":
            resolved = _resolve_dispatch_model("groq", model)
            text, usage = _groq_complete(system, user, resolved, max_tokens)
            return text, usage, "groq", resolved
        if LLM_PROVIDER == "opencode":
            resolved = _resolve_dispatch_model("opencode", model)
            text, usage = _opencode_complete(system, user, resolved, max_tokens)
            return text, usage, "opencode", resolved
        if LLM_PROVIDER == "openai":
            resolved = _resolve_dispatch_model("openai", model)
            text, usage = _openai_complete(system, user, resolved, max_tokens)
            return text, usage, "openai", resolved
        resolved = _resolve_dispatch_model("anthropic", model)
        text, usage = _anthropic_complete(system, user, resolved, max_tokens)
        return text, usage, "anthropic", resolved
    except httpx.HTTPStatusError as e:
        detail = ""
        try:
            detail = e.response.text[:500]
        except Exception:
            pass
        if (
            LLM_PROVIDER == "anthropic"
            and LLM_FALLBACK_TO_OPENCODE
            and _opencode_key()
            and _anthropic_quota_exceeded(e.response.status_code, detail)
        ):
            fb = _opencode_fallback_model(model)
            print(
                f"  [llm] Anthropic limit hit — falling back to OpenCode ({fb})",
                flush=True,
            )
            text, usage = _opencode_complete(system, user, fb, max_tokens)
            return text, usage, "opencode", fb
        return (
            json.dumps(
                {
                    "error": f"http_{e.response.status_code}: {detail or str(e)}",
                    "status": "llm_failed",
                }
            ),
            dict(_ZERO_USAGE),
            LLM_PROVIDER,
            _resolve_dispatch_model(LLM_PROVIDER, model),
        )
    except httpx.HTTPError as e:
        return (
            json.dumps({"error": str(e), "status": "llm_failed"}),
            dict(_ZERO_USAGE),
            LLM_PROVIDER,
            _resolve_dispatch_model(LLM_PROVIDER, model),
        )


def llm_complete(
    system: str,
    user: str,
    *,
    model: str | None = None,
    max_tokens: int = 8192,
) -> str:
    text, _usage, _provider, _model = _dispatch(system, user, model=model, max_tokens=max_tokens)
    return text


def llm_complete_with_usage(
    system: str,
    user: str,
    *,
    model: str | None = None,
    max_tokens: int = 8192,
) -> tuple[str, dict[str, Any]]:
    """Like llm_complete, but also returns {"input_tokens", "output_tokens",
    "provider", "model"} — best-effort (zeros if the provider/mock path didn't
    return usage). Used by call sites that want real cost tracking instead of
    the flat per-call token estimates in pipeline/llm_meter.py.
    """
    text, usage, provider, resolved_model = _dispatch(system, user, model=model, max_tokens=max_tokens)
    return text, {**usage, "provider": provider, "model": resolved_model}


def opencode_model_chain(primary: str | None = None) -> list[str]:
    """Ordered OpenCode models to try (primary first, then fallbacks)."""
    chain: list[str] = []
    if primary:
        chain.append(primary)
    for m in OPENCODE_FALLBACK_MODELS:
        if m not in chain:
            chain.append(m)
    if not chain:
        chain = [OPENCODE_MODEL]
    return chain


def opencode_models_for_attempt(attempt: int, primary: str | None = None) -> str:
    """Rotate OpenCode models when the primary returns 500 or empty."""
    chain = opencode_model_chain(primary)
    return chain[attempt % len(chain)]


def _parse_json_response(raw: str) -> Any:
    if not raw or not str(raw).strip():
        return {"error": "empty_response", "status": "llm_failed"}
    raw_str = str(raw)
    try:
        parsed = _extract_json(raw_str)
        if isinstance(parsed, dict) and parsed.get("status") == "placeholder":
            return {"error": "placeholder", "status": "llm_failed"}
        return parsed
    except (ValueError, json.JSONDecodeError) as e:
        salvaged = salvage_compiler_payload(raw_str)
        if salvaged.get("scope"):
            salvaged["status"] = "llm_partial"
            salvaged["_salvaged"] = True
            return salvaged
        return {
            "error": f"json_parse_failed: {e}",
            "status": "llm_failed",
            "raw_preview": raw_str[:800],
            "_raw_text": raw_str,
        }


def llm_complete_json(
    system: str,
    user: str,
    *,
    model: str | None = None,
    max_tokens: int = 4096,
) -> Any:
    raw = llm_complete(
        system + "\nRespond with valid JSON only. No markdown fences.",
        user,
        model=model,
        max_tokens=max_tokens,
    )
    return _parse_json_response(raw)


def llm_complete_json_with_usage(
    system: str,
    user: str,
    *,
    model: str | None = None,
    max_tokens: int = 4096,
) -> tuple[Any, dict[str, Any]]:
    """Like llm_complete_json, but also returns the usage dict from
    llm_complete_with_usage — see that function's docstring."""
    raw, usage = llm_complete_with_usage(
        system + "\nRespond with valid JSON only. No markdown fences.",
        user,
        model=model,
        max_tokens=max_tokens,
    )
    return _parse_json_response(raw), usage
