"""Custom CSS for the Vendor Intelligence Streamlit app."""

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&family=JetBrains+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', system-ui, sans-serif;
}

.hero {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 45%, #0d9488 100%);
    border-radius: 16px;
    padding: 2rem 2.25rem;
    margin-bottom: 1.5rem;
    color: #f8fafc;
    box-shadow: 0 20px 40px rgba(15, 23, 42, 0.25);
}
.hero h1 {
    font-size: 2rem;
    font-weight: 700;
    margin: 0 0 0.5rem 0;
    letter-spacing: -0.02em;
}
.hero p {
    margin: 0;
    opacity: 0.9;
    font-size: 1.05rem;
    line-height: 1.5;
}
.metric-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1.1rem 1.25rem;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.06);
}
.metric-card .label {
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: #64748b;
    font-weight: 600;
}
.metric-card .value {
    font-size: 1.75rem;
    font-weight: 700;
    color: #0f172a;
    margin-top: 0.25rem;
}
.phase-pill {
    display: inline-block;
    background: #f0fdfa;
    color: #0f766e;
    border: 1px solid #99f6e4;
    border-radius: 999px;
    padding: 0.25rem 0.75rem;
    font-size: 0.8rem;
    font-weight: 600;
    margin-right: 0.35rem;
    margin-bottom: 0.35rem;
}
.status-ok { color: #059669; font-weight: 600; }
.status-warn { color: #d97706; font-weight: 600; }
.status-fail { color: #dc2626; font-weight: 600; }
.log-box {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.78rem;
    background: #0f172a;
    color: #e2e8f0;
    border-radius: 10px;
    padding: 1rem;
    max-height: 420px;
    overflow-y: auto;
    white-space: pre-wrap;
    line-height: 1.45;
}
.sidebar-brand {
    font-size: 1.1rem;
    font-weight: 700;
    color: #0f172a;
    padding: 0.5rem 0 1rem 0;
    border-bottom: 2px solid #0d9488;
    margin-bottom: 1rem;
}
.result-card {
    border-left: 4px solid #0d9488;
    background: #f8fafc;
    border-radius: 0 10px 10px 0;
    padding: 0.85rem 1rem;
    margin-bottom: 0.5rem;
}
div[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
}
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #0d9488, #0f766e);
    border: none;
    font-weight: 600;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #14b8a6, #0d9488);
    border: none;
}
</style>
"""
