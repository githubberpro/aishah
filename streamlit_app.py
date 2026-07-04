"""Aishah — Streamlit app entry point.

Streamlit Community Cloud (share.streamlit.io) looks for `streamlit_app.py`
by default, so keeping the entry point here makes deployment one click.
"""

import streamlit as st


def greeting(name: str) -> str:
    """Return a friendly greeting. Kept pure so it is easy to unit-test."""
    name = name.strip() or "world"
    return f"Hello, {name}!"


def main() -> None:
    st.set_page_config(page_title="Aishah", page_icon="✨", layout="centered")

    st.title("✨ Aishah")
    st.write("A starter Streamlit app, deployed on Streamlit Community Cloud.")

    name = st.text_input("What's your name?", value="world")
    st.subheader(greeting(name))

    with st.expander("About this app"):
        st.markdown(
            "This is a minimal scaffold. Replace `streamlit_app.py` with your "
            "real app — the deploy and test pipeline will keep working."
        )


if __name__ == "__main__":
    main()
