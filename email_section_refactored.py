# ═══════════════════════════════════════════════════════════════���
#  REFACTORED EMAIL SECTION
#  Replace the email section in page_create_quote() (lines 1368-1499)
#  with this code
# ════════════════════════════════════════════════════════════════

# ── Import at top of file ──
# from email_templates import (
#     create_outlook_email_message,
#     generate_email_html,
#     generate_plain_text_body,
#     generate_outlook_email_table,
#     generate_totals_table,
#     safe_str,
#     format_currency,
#     format_number
# )


# ── Replace this section (lines 1368-1499 in purchase_quote_app.py) ──
def email_section(
    show_actions_for: str,
    q_rows_for_actions: pd.DataFrame,
    vendor_email: str,
    buf_pdf: io.BytesIO
):
    """
    Refactored email submission section with Outlook compatibility.
    
    ✅ Features:
        - Outlook-compatible HTML tables
        - No broken HTML tags
        - Proper character encoding
        - Works in Gmail, Outlook Web, Outlook Classic
    """
    
    # Import the email template functions
    from email_templates import (
        create_outlook_email_message,
        generate_email_html,
        generate_outlook_email_table,
        generate_totals_table,
        safe_str,
        format_currency
    )
    
    st.write("**✉️ Email Submission**")
    
    # ── Calculate totals ──
    total_qty = 0
    subtotal_excl_gst = 0.0
    total_gst = 0.0
    grand_total_before_freight = 0.0
    
    for _, row in q_rows_for_actions.iterrows():
        try:
            qty = float(row.get('Qty', 0) or 0)
            total_qty += qty
        except (ValueError, TypeError):
            qty = 0.0
        
        try:
            price_excl = float(row.get('Price Before GST', 0) or 0)
            price_incl = float(row.get('Price Inc. GST', 0) or 0)
        except (ValueError, TypeError):
            price_excl = 0.0
            price_incl = 0.0
        
        line_excl = qty * price_excl
        line_incl = qty * price_incl
        
        subtotal_excl_gst += line_excl
        total_gst += (line_incl - line_excl)
        grand_total_before_freight += line_incl
    
    try:
        raw_freight = q_rows_for_actions["Freight Charge"].iloc[0] if "Freight Charge" in q_rows_for_actions.columns else 0
        freight_charge = float(raw_freight) if pd.notnull(raw_freight) else 0.0
    except (ValueError, TypeError):
        freight_charge = 0.0
    
    grand_total = grand_total_before_freight + freight_charge
    
    # ── Editable email content ──
    def_greeting = f"""Dear Sir/Madam,

Greetings from {safe_str(vendor_email.split('@')[0] if '@' in vendor_email else 'our company')}.

Please find below the Purchase Quote details for your reference. We would appreciate your prompt review and confirmation."""
    
    def_closing = f"""Total Qty: {int(total_qty)}
Total Amount (Incl. GST): {format_currency(grand_total)}

Kindly review and provide your confirmation at the earliest. Please feel free to contact us for any clarification or modifications.

Thank you for your business!"""
    
    c_email1, c_email2 = st.columns(2)
    
    with c_email1:
        email_greeting = st.text_area(
            "📝 Email Greeting & Intro",
            value=def_greeting,
            height=100,
            help="Customize the email opening message"
        )
    
    with c_email2:
        email_closing = st.text_area(
            "📝 Email Closing & Signature",
            value=def_closing,
            height=100,
            help="Customize the email closing message"
        )
    
    # ── Email parameters ──
    subject = f"Purchase Quote Submission - {show_actions_for}"
    cc_emails = "purchase@supremeindia.com, mis3@supremeindia.com"
    to_email = vendor_email if vendor_email and safe_str(vendor_email) != "nan" else ""
    
    if not to_email:
        st.warning("⚠️ No vendor email address found. Please update vendor details with email address.")
        return
    
    # ══════════════════════════════════════════════════════════════
    #  OPTION 1: Download .EML file (opens in Outlook)
    # ══════════════════════════════════════════════════════════════
    
    st.markdown("##### 📎 Download Email Draft")
    
    # Create email message with Outlook-compatible HTML
    msg = create_outlook_email_message(
        subject=subject,
        to_email=to_email,
        cc_emails=cc_emails,
        greeting=email_greeting,
        quote_data=q_rows_for_actions,
        subtotal_excl_gst=subtotal_excl_gst,
        total_gst=total_gst,
        freight_charge=freight_charge,
        grand_total=grand_total,
        closing=email_closing,
        pdf_attachment=buf_pdf,
        pdf_filename=f"PO_{show_actions_for}.pdf"
    )
    
    # Download .EML file
    eml_bytes = msg.as_bytes()
    st.download_button(
        label="⬇️ Download .EML (Double-click to open in Outlook)",
        data=eml_bytes,
        file_name=f"PO_{show_actions_for}_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.eml",
        mime="message/rfc822",
        use_container_width=True
    )
    
    st.caption("✓ The .EML file includes:")
    st.caption("  • Recipient email (To, Cc)")
    st.caption("  • Subject line")
    st.caption("  • Formatted HTML table (works in Outlook Classic)")
    st.caption("  • Plain text fallback")
    st.caption("  • PDF attachment (PO document)")
    
    st.markdown("---")
    
    # ══════════════════════════════════════════════════════════════
    #  OPTION 2: Preview HTML in Streamlit
    # ══════════════════════════════════════════════════════════════
    
    with st.expander("👁️ Preview Email in Browser"):
        st.write("**Email Preview** (this is how it will appear in Outlook):")
        
        # Generate and display HTML
        html_preview = generate_email_html(
            greeting=email_greeting,
            quote_data=q_rows_for_actions,
            subtotal_excl_gst=subtotal_excl_gst,
            total_gst=total_gst,
            freight_charge=freight_charge,
            grand_total=grand_total,
            closing=email_closing
        )
        
        st.markdown(html_preview, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # ══════════════════════════════════════════════════════════════
    #  OPTION 3: Mailto Link (opens email client)
    # ══════════════════════════════════════════════════════════════
    
    st.markdown("##### 🖱️ Or Open Email Client Directly")
    
    # Generate plain text body for mailto
    from email_templates import generate_plain_text_body
    
    text_body = generate_plain_text_body(
        greeting=email_greeting,
        quote_data=q_rows_for_actions,
        subtotal_excl_gst=subtotal_excl_gst,
        total_gst=total_gst,
        freight_charge=freight_charge,
        grand_total=grand_total,
        closing=email_closing
    )
    
    # Build mailto URL
    mailto_params = {
        "subject": subject,
        "cc": cc_emails,
        "body": text_body
    }
    mailto_url = f"mailto:{urllib.parse.quote(to_email)}?{urllib.parse.urlencode(mailto_params, quote_via=urllib.parse.quote)}"
    
    st.markdown(f'''
    <a href="{mailto_url}" style="display:inline-block;padding:12px 24px;background:#059669;color:white;text-decoration:none;border-radius:8px;font-weight:bold;font-size:15px;">
        ✉️ Open Outlook / Default Email Client
    </a>
    ''', unsafe_allow_html=True)
    
    st.caption("Click to open your default email client with pre-filled recipient, CC, and subject.")
    st.caption("**Note:** Attach the PDF manually after downloading it above.")


# ════════════════════════════════════════════════════════════════
#  HOW TO USE IN page_create_quote()
# ════════════════════════════════════════════════════════════���═══

# Replace lines 1368-1499 with this call:

# if show_actions_for:
#     st.markdown("---")
#     st.subheader(f"Actions for Quote: {show_actions_for}")
#     
#     q_rows_for_actions = quotes_df[quotes_df["Quote No"] == show_actions_for]
#     
#     c_dl, c_mail = st.columns([1.5, 2.5])
#     
#     with c_dl:
#         st.write("**📥 Export PO**")
#         buf_xl = export_po_excel(show_actions_for, quotes_df)
#         if buf_xl:
#             st.download_button(
#                 "Excel PO",
#                 buf_xl,
#                 file_name=f"PO_{show_actions_for}.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                 use_container_width=True
#             )
#         buf_pdf = export_po_pdf(show_actions_for, quotes_df)
#         if buf_pdf:
#             st.download_button(
#                 "PDF PO",
#                 buf_pdf,
#                 file_name=f"PO_{show_actions_for}.pdf",
#                 mime="application/pdf",
#                 use_container_width=True
#             )
#     
#     with c_mail:
#         email_section(
#             show_actions_for=show_actions_for,
#             q_rows_for_actions=q_rows_for_actions,
#             vendor_email=q_rows_for_actions["Vendor Email"].iloc[0] if not q_rows_for_actions.empty else "",
#             buf_pdf=buf_pdf
#         )
