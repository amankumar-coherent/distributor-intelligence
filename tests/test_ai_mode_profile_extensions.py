"""Automation-profile extension persistence.

These tests pin the behaviour that made hand-adding an extension possible.
Before this, two things defeated it: `--disable-extensions-except` hid the
profile's own extensions (so chrome://extensions looked EMPTY and "Load
unpacked" appeared to do nothing), and an automatic profile rebuild deleted
whatever the user had installed. Both are regressions worth catching.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from vendor_intel.scraping import google_ai_mode as aim


@pytest.fixture
def profile(tmp_path, monkeypatch):
    """Point the module at a throwaway profile dir."""
    target = tmp_path / "ai_mode_chrome_profile"
    monkeypatch.setenv("GOOGLE_AI_MODE_PROFILE_DIR", str(target))
    return target


def _make_extension(root, name="captcha-raptor", nested=False):
    folder = root / name
    manifest_dir = folder / "extension" if nested else folder
    manifest_dir.mkdir(parents=True, exist_ok=True)
    (manifest_dir / "manifest.json").write_text(
        json.dumps({"name": name, "version": "1.0", "manifest_version": 3}),
        encoding="utf-8",
    )
    (manifest_dir / "background.js").write_text("// worker", encoding="utf-8")
    return folder


# --- defaults -------------------------------------------------------------


def test_keep_profile_defaults_on(monkeypatch):
    """The profile must be preserved by default, otherwise a manual
    install silently disappears on the next run."""
    monkeypatch.delenv("GOOGLE_AI_MODE_KEEP_PROFILE", raising=False)
    assert aim.keep_profile() is True


def test_allow_all_extensions_defaults_on(monkeypatch):
    """Must default on: with --disable-extensions-except set, the profile's
    own extensions are suppressed and chrome://extensions shows nothing."""
    monkeypatch.delenv("GOOGLE_AI_MODE_ALLOW_ALL_EXTENSIONS", raising=False)
    assert aim.allow_all_extensions() is True


@pytest.mark.parametrize("value", ["false", "0", "no", "off"])
def test_flags_can_be_disabled(monkeypatch, value):
    monkeypatch.setenv("GOOGLE_AI_MODE_KEEP_PROFILE", value)
    monkeypatch.setenv("GOOGLE_AI_MODE_ALLOW_ALL_EXTENSIONS", value)
    assert aim.keep_profile() is False
    assert aim.allow_all_extensions() is False


# --- install_extension ----------------------------------------------------


def test_install_extension_copies_into_profile(tmp_path, profile):
    src = _make_extension(tmp_path / "downloads")

    assert aim.install_extension(str(src)) is True

    staged = aim._staging_root() / "captcha-raptor"
    assert (staged / "manifest.json").is_file()
    # The copy must be self-contained: referenced files come too, or the
    # extension loads broken.
    assert (staged / "background.js").is_file()


def test_install_extension_finds_nested_manifest(tmp_path, profile):
    """captcha-raptor keeps its manifest in extension/, and selecting the
    repo root is the most common mistake — auto-correct it."""
    src = _make_extension(tmp_path / "downloads", nested=True)

    assert aim.install_extension(str(src)) is True
    manifests = list(aim._staging_root().rglob("manifest.json"))
    assert len(manifests) == 1


def test_install_names_stage_after_extension_not_generic_folder(tmp_path, profile):
    """Selecting `captcha-raptor\\extension` must stage as "captcha-raptor",
    not "extension" — otherwise two extensions with the same common layout
    collide and the listing is unreadable."""
    src = _make_extension(tmp_path / "downloads", name="captcha-raptor", nested=True)

    assert aim.install_extension(str(src / "extension")) is True

    staged = aim.staged_extension_paths()
    assert len(staged) == 1
    assert staged[0].endswith("captcha-raptor")


def test_install_two_extensions_with_same_layout_do_not_collide(tmp_path, profile):
    """hektcaptcha keeps its manifest in public/, lets-solve-captcha in
    src/ — staging both must not overwrite one another."""
    a = _make_extension(tmp_path / "downloads", name="captcha-raptor", nested=True)
    b = tmp_path / "downloads" / "hektcaptcha" / "public"
    b.mkdir(parents=True)
    (b / "manifest.json").write_text('{"name":"hekt"}', encoding="utf-8")

    assert aim.install_extension(str(a / "extension")) is True
    assert aim.install_extension(str(b)) is True

    staged = aim.staged_extension_paths()
    assert len(staged) == 2
    names = sorted(p.rsplit("\\", 1)[-1].rsplit("/", 1)[-1] for p in staged)
    assert names == ["captcha-raptor", "hektcaptcha"]


def test_install_extension_rejects_folder_without_manifest(tmp_path, profile):
    bad = tmp_path / "not-an-extension"
    bad.mkdir()
    assert aim.install_extension(str(bad)) is False


def test_install_extension_strips_quotes(tmp_path, profile):
    """Windows "Copy as path" wraps the path in double quotes."""
    src = _make_extension(tmp_path / "downloads")
    assert aim.install_extension(f'"{src}"') is True


def test_install_extension_is_idempotent(tmp_path, profile):
    src = _make_extension(tmp_path / "downloads")
    assert aim.install_extension(str(src)) is True
    assert aim.install_extension(str(src)) is True
    assert len(aim.staged_extension_paths()) == 1


# --- identity / dedupe ----------------------------------------------------


def test_extension_identity_uses_manifest_name(tmp_path):
    """Two paths to the SAME extension must resolve to one identity, so the
    launcher doesn't register it twice (two solver workers racing on one
    CAPTCHA)."""
    original = _make_extension(tmp_path / "downloads", name="captcha-raptor")
    copy = tmp_path / "profile" / "_unpacked_extensions" / "renamed-folder"
    copy.mkdir(parents=True)
    (copy / "manifest.json").write_text(
        json.dumps({"name": "captcha-raptor", "version": "1.0"}), encoding="utf-8"
    )

    assert aim._extension_identity(original) == aim._extension_identity(copy)


def test_extension_identity_distinguishes_different_extensions(tmp_path):
    a = _make_extension(tmp_path / "d", name="captcha-raptor")
    b = _make_extension(tmp_path / "d", name="hektcaptcha")
    assert aim._extension_identity(a) != aim._extension_identity(b)


def test_extension_identity_falls_back_to_folder_name(tmp_path):
    """An unreadable manifest must not crash the launch path."""
    folder = tmp_path / "MyExt"
    folder.mkdir()
    (folder / "manifest.json").write_text("{broken", encoding="utf-8")
    assert aim._extension_identity(folder) == "myext"


# --- staged_extension_paths ----------------------------------------------


def test_staged_paths_empty_when_nothing_installed(profile):
    assert aim.staged_extension_paths() == []


def test_staged_paths_lists_installed_extension(tmp_path, profile):
    src = _make_extension(tmp_path / "downloads")
    aim.install_extension(str(src))

    staged = aim.staged_extension_paths()
    assert len(staged) == 1
    assert "captcha-raptor" in staged[0]


def test_staged_paths_skips_folders_without_manifest(tmp_path, profile):
    """A half-copied or unrelated folder must not be handed to Chrome as an
    extension — Chrome fails the whole --load-extension list if one entry
    is invalid, which would silently disable the good ones too."""
    src = _make_extension(tmp_path / "downloads")
    aim.install_extension(str(src))
    junk = profile / "_unpacked_extensions" / "junk"
    junk.mkdir(parents=True)

    staged = aim.staged_extension_paths()
    assert len(staged) == 1
    assert "junk" not in staged[0]


# --- rebuild preserves installs ------------------------------------------


def test_reset_profile_preserves_staged_extensions(tmp_path, profile):
    """The whole point: clearing browsing state must not cost the user
    their extensions. The old code deleted the profile wholesale."""
    src = _make_extension(tmp_path / "downloads")
    aim.install_extension(str(src))
    # Simulate real browsing state that SHOULD be cleared.
    default = profile / "Default"
    default.mkdir(parents=True, exist_ok=True)
    (default / "Cookies").write_text("flagged-session", encoding="utf-8")

    aim._rebuild_profile_preserving_staged(profile)

    assert not (default / "Cookies").exists(), "browsing state should be cleared"
    assert aim.staged_extension_paths(), "staged extension must survive"
    assert (aim._staging_root() / "captcha-raptor" / "manifest.json").is_file()


def test_staging_lives_outside_the_profile(tmp_path, profile):
    """Staging inside the profile caused two separate bugs: a reset had to
    move it out and back (recreating the profile dir, which defeated the
    `first_run` check), and Chrome/Chromium use different profiles, so a
    copy staged for one build was invisible to the other."""
    src = _make_extension(tmp_path / "downloads")
    aim.install_extension(str(src))

    staged = Path(aim.staged_extension_paths()[0])
    assert profile.resolve() not in staged.resolve().parents


def test_reset_profile_handles_no_staged_extensions(tmp_path, profile):
    default = profile / "Default"
    default.mkdir(parents=True, exist_ok=True)
    (default / "Cookies").write_text("x", encoding="utf-8")

    aim._rebuild_profile_preserving_staged(profile)  # must not raise

    assert not (default / "Cookies").exists()


# --- first_run detection --------------------------------------------------


def test_reset_leaves_no_chrome_state_so_relaunch_is_first_run(tmp_path, profile):
    """`first_run` keys off Chrome's own state (Default/), not merely the
    profile folder existing. A reset must leave no Default/ behind, so the
    next launch re-seeds and is treated as a first run — while the staged
    extension (stored outside) still survives."""
    src = _make_extension(tmp_path / "downloads")
    aim.install_extension(str(src))
    (profile / "Default").mkdir(parents=True, exist_ok=True)

    aim._rebuild_profile_preserving_staged(profile)

    assert not (profile / "Default").exists(), (
        "Chrome state must be gone so the relaunch counts as first_run"
    )
    assert aim.staged_extension_paths(), "staged extension must survive a reset"


# --- extension IDs / managed-Chrome fallback ------------------------------


def test_unpacked_extension_id_matches_chrome(monkeypatch):
    """Pins the derivation against a REAL observed value.

    Ground truth, read from chrome://extensions-internals after Chromium
    loaded this exact staged path: the extension appeared as
    ejbanjngfemjkclbldgccmkncjjmllln. Chrome derives the ID from a SHA-256
    of the absolute path (UTF-16LE on Windows), first 16 bytes, nibbles
    mapped 0-f -> a-p. If this ever drifts, worker attribution silently
    breaks and a non-running solver looks healthy again.
    """
    import os as _os

    if _os.name != "nt":
        pytest.skip("path-hash derivation pinned on the Windows observation")

    path = (
        r"C:\Users\aman.kumar\Downloads\vendor-intel\data"
        r"\ai_mode_chrome_profile\_unpacked_extensions\captcha-raptor"
    )
    monkeypatch.setattr(aim.Path, "resolve", lambda self: self, raising=False)
    assert aim.unpacked_extension_id(path) == "ejbanjngfemjkclbldgccmkncjjmllln"


def test_unpacked_extension_id_is_stable_and_path_specific(tmp_path):
    a = tmp_path / "one"
    b = tmp_path / "two"
    a.mkdir()
    b.mkdir()
    assert aim.unpacked_extension_id(a) == aim.unpacked_extension_id(a)
    assert aim.unpacked_extension_id(a) != aim.unpacked_extension_id(b)


def test_unpacked_extension_id_shape():
    """32 chars, all in a-p — anything else means the nibble mapping broke."""
    ext_id = aim.unpacked_extension_id(r"C:\some\path")
    assert len(ext_id) == 32
    assert all("a" <= c <= "p" for c in ext_id)


def test_extensions_require_chromium_defaults_on(monkeypatch):
    """Measured on Chrome 152 with an enterprise ExtensionInstallForcelist
    policy set: real Chrome does NOT load a command-line unpacked
    extension (confirmed via chrome://extensions-internals, both with and
    without --disable-extensions-except), while bundled Chromium does. So
    honouring an extension request means switching build by default."""
    monkeypatch.delenv("GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM", raising=False)
    assert aim.extensions_require_chromium() is True


def test_extensions_require_chromium_can_be_disabled(monkeypatch):
    monkeypatch.setenv("GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM", "false")
    assert aim.extensions_require_chromium() is False


def test_chromium_gets_a_separate_profile_dir(profile):
    """Chrome and Chromium must never share a profile directory — they are
    different builds with incompatible state and cookie encryption."""
    assert aim._profile_dir_for("chrome") != aim._profile_dir_for(None)
    assert aim._profile_dir_for(None).endswith("_chromium")


def test_active_profile_dir_follows_the_chromium_switch(tmp_path, monkeypatch, profile):
    """--list/--reset-profile must report the directory the pipeline will
    actually open, not the Chrome-channel default."""
    monkeypatch.setenv("GOOGLE_AI_MODE_BROWSER", "chrome")
    monkeypatch.setenv("GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM", "true")
    src = _make_extension(tmp_path / "downloads")
    aim.install_extension(str(src))

    assert aim.active_profile_dir().endswith("_chromium")

    monkeypatch.setenv("GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM", "false")
    assert not aim.active_profile_dir().endswith("_chromium")


def test_active_profile_dir_unchanged_without_extensions(monkeypatch, profile):
    monkeypatch.setenv("GOOGLE_AI_MODE_BROWSER", "chrome")
    monkeypatch.delenv("GOOGLE_AI_MODE_EXTENSIONS", raising=False)
    assert aim.active_profile_dir() == aim._profile_dir()


def _ctx_with_workers(*urls):
    return type(
        "Ctx",
        (),
        {
            "service_workers": [type("W", (), {"url": u})() for u in urls],
            "background_pages": [],
        },
    )()


def test_component_extension_workers_are_not_counted(tmp_path):
    """Chrome always runs pdf/hangouts/network-speech workers. Counting
    them reported "extension active" while the solver was never loaded —
    the exact false positive that hid this problem for several rounds."""
    ext = _make_extension(tmp_path, name="captcha-raptor")
    ctx = _ctx_with_workers(
        f"chrome-extension://{aim._COMPONENT_EXTENSION_IDS[0]}/service_worker.js"
    )

    assert aim._own_extension_workers(ctx, [str(ext)]) == []


def test_own_extension_worker_is_detected(tmp_path):
    """The positive case: a worker whose ID matches the loaded path counts."""
    ext = _make_extension(tmp_path, name="captcha-raptor")
    ext_id = aim.unpacked_extension_id(ext)
    ctx = _ctx_with_workers(
        f"chrome-extension://{aim._COMPONENT_EXTENSION_IDS[0]}/service_worker.js",
        f"chrome-extension://{ext_id}/background.js",
    )

    ours = aim._own_extension_workers(ctx, [str(ext)])
    assert len(ours) == 1
    assert ext_id in ours[0]


def test_own_extension_workers_uses_paths_actually_loaded(tmp_path):
    """Attribution must key off the paths handed to Chrome, not a
    re-derivation from env/staging. The launch path dedupes those two
    sources, so re-deriving could yield a different set than Chrome got
    and silently mis-attribute the worker."""
    loaded = _make_extension(tmp_path / "staged", name="captcha-raptor")
    other = _make_extension(tmp_path / "elsewhere", name="captcha-raptor")

    ctx = _ctx_with_workers(
        f"chrome-extension://{aim.unpacked_extension_id(loaded)}/background.js"
    )

    assert aim._own_extension_workers(ctx, [str(loaded)])
    # Same manifest name, different path => different ID => not a match.
    assert aim._own_extension_workers(ctx, [str(other)]) == []


# --- _installed_extension_ids --------------------------------------------


def _write_prefs(profile, settings):
    default = profile / "Default"
    default.mkdir(parents=True, exist_ok=True)
    (default / "Preferences").write_text(
        json.dumps({"extensions": {"settings": settings}}), encoding="utf-8"
    )


def test_installed_ids_empty_without_preferences(profile):
    assert aim._installed_extension_ids(profile) == []


def test_installed_ids_reports_user_installed(profile):
    _write_prefs(profile, {"abcdefghijklmnop": {"location": 4, "state": 1}})
    assert aim._installed_extension_ids(profile) == ["abcdefghijklmnop"]


def test_installed_ids_filters_component_extensions(profile):
    """Chrome ships built-in component extensions; counting them would
    make an empty profile look like it has extensions installed."""
    _write_prefs(
        profile,
        {
            "builtin000000000": {"location": 5, "state": 1},
            "policy0000000000": {"location": 10, "state": 1},
            "userinstalled000": {"location": 4, "state": 1},
        },
    )
    assert aim._installed_extension_ids(profile) == ["userinstalled000"]


def test_installed_ids_skips_disabled(profile):
    """A disabled extension can't solve a CAPTCHA, so reporting it as
    present would be misleading."""
    _write_prefs(profile, {"disabled00000000": {"location": 4, "state": 0}})
    assert aim._installed_extension_ids(profile) == []


def test_installed_ids_survives_corrupt_preferences(profile):
    """Chrome can leave Preferences mid-write; a crash here would take
    down the whole pipeline launch for a purely informational read."""
    default = profile / "Default"
    default.mkdir(parents=True, exist_ok=True)
    (default / "Preferences").write_text("{not valid json", encoding="utf-8")

    assert aim._installed_extension_ids(profile) == []
