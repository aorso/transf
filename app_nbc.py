"""
NBC TermSheet Transformer — Streamlit App
Applies the standard NBC transformation process to any uploaded .docx term sheet
and provides two output files for download: mark-up and final.
"""

import io
import tempfile
from pathlib import Path

import streamlit as st

from ts_editor_D import TermSheetTransformer


# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="NBC TermSheet Transformer",
    page_icon="📄",
    layout="centered",
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def run_nbc_transformation(input_path: Path) -> tuple[bytes, bytes]:
    """
    Runs the full NBC transformation process on the given .docx file.
    Returns (markup_bytes, final_bytes).
    """
    transformer = TermSheetTransformer(str(input_path))

    # 3.1 — Word replacements
    transformer.replace_word("Notes", "Securities")
    transformer.replace_word("notes", "securities")
    transformer.replace_word("Note", "Security")
    transformer.replace_word("note", "security")

    transformer.replace_word("Certificates", "Securities")
    transformer.replace_word("certificates", "securities")
    transformer.replace_word("Certificate", "Security")
    transformer.replace_word("certificate", "security")

    transformer.replace_word("Status of the Notes and Guarantees", "Status of the Securities")
    transformer.replace_word("Status of the Certificates and Guarantees", "Status of the Securities")
    transformer.replace_word("Status of the Securities and Guarantees", "Status of the Securities")

    # 3.2 — Section descriptions
    transformer.set_section(title="Issuer", description="National Bank of Canada")
    transformer.set_section(
        title="Issuer's Domicile",
        description="800 Saint-Jacques Street, Montréal, Québec, Canada H3C 1A3",
    )
    transformer.set_section(title="Calculation Agent", description="National Bank of Canada")
    transformer.set_section(title="Issue Type", description="Structured Securities")
    transformer.set_section(
        title="Codes",
        description="ISIN: TBD \n- Common: TBD \n- CFI: TBD \n- FISH: TBD",
        red_highlight_in_final=True,
    )
    transformer.set_section(title="Common Depositary", description="Citibank Europe PLC")

    # 3.3 — Section removals
    for section in [
        "FCNA",
        "Issuer's Prudential Supervision",
        "Guarantor",
        "Guarantor's Domicile",
        "Guarantor's Prudential Supervision",
        "Principal Security Agent",
        "Principal Security Agent's Domicile",
        "Calculation Agent's Domicile",
        "Seniority",
        "Security",
        "Issuer's Web Page / Publication",
        "Common Depositary's Domicile",
    ]:
        transformer.remove_section(section)

    # 3.4 — Logo
    logo_path = Path(__file__).parent / "logo_nbc.png"
    if logo_path.exists():
        transformer.add_logo(logo_path=str(logo_path), width_inches=1.6)

    # 3.5 — Disclaimer sections
    transformer.set_disclaimer_section(
        title="New Title",
        content="Notice importante",
        red_highlight_in_final=True,
    )

    # 3.6 — Remove disclaimer sections
    transformer.remove_disclaimer_section("IMPORTANT INFORMATION")

    # Section order
    transformer.set_section_order([
        "Issue Type",
        "Issuer",
        "Series Number",
        "Bloomberg",
        "Bail-In",
        "Issuance Program",
        "Long-Term Non Bail-inable Debt Issuer Ratings",
        "Dealer",
        "Currency",
        "Public Offer",
        "Listing and Trading",
        "Calculation Agent",
        "Issue Amount",
        "Number of Securities",
        "Specified Denomination (D )",
        "Issue Price per Security",
        "Listing",
        "Minimum Trading Size",
        "Trade Date",
        "Strike Date",
        "Issue Date",
        "Redemption Valuation Date",
        "Maturity Date",
        "Underlying Index",
        "Administration",
        "Registered",
        "Coupon",
        "Conditional Coupon",
        "Final Redemption",
        "Where",
        "Business Day Convention",
        "Payment Business Days",
        "Governing Law",
        "Status of the Securities",
        "Form of Global Security Codes",
        "Reuters RIC Code",
        "Common Depositary",
        "Secondary Trading",
        "Initial Settlement",
        "Fees",
        "D. Restrictions",
        "No offering information",
        "Prohibition of Sales to EEA Retail Investors",
        "Prohibition of Sales to UK Retail Investors",
        "Target market",
    ])

    # Export to a temporary directory and read bytes back
    with tempfile.TemporaryDirectory() as tmp_dir:
        transformer.execute_and_export(
            output_dir=tmp_dir,
            markup_docx="markup.docx",
            final_docx="final.docx",
            final_pdf="final.pdf",  # won't be used but required by the API
        )
        markup_bytes = (Path(tmp_dir) / "markup.docx").read_bytes()
        final_bytes = (Path(tmp_dir) / "final.docx").read_bytes()

    return markup_bytes, final_bytes


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.title("📄 NBC TermSheet Transformer")
st.markdown(
    "Upload a **.docx** term sheet, run the standard NBC transformation, "
    "and download the two output documents."
)

st.divider()

# --- Upload ---
uploaded_file = st.file_uploader(
    "Drop your term sheet here or click to browse",
    type=["docx"],
    help="Only .docx files are accepted.",
)

if uploaded_file is not None:
    st.success(f"File loaded: **{uploaded_file.name}**")

    st.divider()

    # --- Run button ---
    if st.button("🚀 Run NBC Transformation", use_container_width=True, type="primary"):
        with st.spinner("Applying NBC transformations…"):
            try:
                # Write upload to a temp file so TermSheetTransformer can read it
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_in:
                    tmp_in.write(uploaded_file.getvalue())
                    tmp_in_path = Path(tmp_in.name)

                markup_bytes, final_bytes = run_nbc_transformation(tmp_in_path)
                tmp_in_path.unlink(missing_ok=True)

                st.success("Transformation complete!")
                st.divider()

                # --- Downloads ---
                st.subheader("⬇️ Download results")

                col1, col2 = st.columns(2)

                stem = Path(uploaded_file.name).stem

                with col1:
                    st.download_button(
                        label="📝 Mark-up version",
                        data=markup_bytes,
                        file_name=f"{stem}_markup.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                    st.caption("Changes highlighted — yellow additions, strikethrough deletions")

                with col2:
                    st.download_button(
                        label="✅ Final version",
                        data=final_bytes,
                        file_name=f"{stem}_final.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                    st.caption("Clean final document, no markup")

            except Exception as exc:
                tmp_in_path.unlink(missing_ok=True)
                st.error(f"An error occurred during transformation:\n\n`{exc}`")

st.divider()
st.caption("NBC TermSheet Transformer · Powered by python-docx")
