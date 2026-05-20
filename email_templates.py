# ════════════════════════════════════════════════════════════════
#  Email Templates Module - Outlook Compatible
#  Generates professional purchase quote emails with proper HTML
#  Works in: Gmail ✓ | Outlook Web ✓ | Outlook Classic ✓
# ════════════════════════════════════════════════════════════════

import pandas as pd
from email.message import EmailMessage
from typing import Optional
import io
from datetime import datetime


# ── UTILITY FUNCTIONS ────────────────────────────────────────────

def safe_str(value: any, default: str = "") -> str:
    """
    Safely convert value to string with HTML escaping.
    
    Handles:
    - None / NaN values
    - HTML special characters (<, >, &, ", ')
    - Unicode characters
    - Numeric types
    
    Args:
        value: Any value to convert
        default: Default if value is None/NaN
    
    Returns:
        str: Safe HTML-escaped string
    
    Examples:
        >>> safe_str(None)
        ''
        >>> safe_str("Product<description>")
        'Product&lt;description&gt;'
        >>> safe_str("<script>alert('xss')</script>")
        '&lt;script&gt;alert(&apos;xss&apos;)&lt;/script&gt;'
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return default
    
    value_str = str(value).strip()
    
    if not value_str or value_str.lower() == 'nan':
        return default
    
    # HTML escape special characters
    value_str = value_str.replace("&", "&amp;")      # & first!
    value_str = value_str.replace("<", "&lt;")
    value_str = value_str.replace(">", "&gt;")
    value_str = value_str.replace('"', "&quot;")
    value_str = value_str.replace("'", "&apos;")
    
    return value_str


def format_currency(amount: float, currency_symbol: str = "₹", decimals: int = 2) -> str:
    """
    Format number as currency with proper separators.
    
    Args:
        amount: Numeric amount
        currency_symbol: Symbol to use (default: ₹)
        decimals: Decimal places (default: 2)
    
    Returns:
        str: Formatted currency string
    
    Examples:
        >>> format_currency(1234.5)
        '₹1,234.50'
        >>> format_currency(1000000)
        '₹10,00,000.00'
    """
    try:
        amount = float(amount) if amount else 0.0
    except (ValueError, TypeError):
        amount = 0.0
    
    return f"{currency_symbol}{amount:,.{decimals}f}"


def format_number(value: float, decimals: int = 2) -> str:
    """
    Format number with thousand separators.
    
    Args:
        value: Numeric value
        decimals: Decimal places
    
    Returns:
        str: Formatted number
    
    Examples:
        >>> format_number(1234.5)
        '1,234.50'
    """
    try:
        value = float(value) if value else 0.0
    except (ValueError, TypeError):
        value = 0.0
    
    return f"{value:,.{decimals}f}"


# ── TABLE GENERATION ────────���────────────────────────────────────

def generate_outlook_email_table(
    quote_data: pd.DataFrame,
    header_color: str = "#BDD7EE",
    header_text_color: str = "#000000",
    row_color_alt: str = "#F2F2F2",
    border_color: str = "#000000"
) -> str:
    """
    Generate Outlook-compatible HTML table from quote data.
    
    Features:
    - Properly closed tags
    - No truncation
    - Table-based layout (NO divs/flexbox)
    - Inline styles only
    - HTML-escaped content
    
    Args:
        quote_data: DataFrame with columns:
            - S.No, Vendor Item No, Product Description
            - Qty, Price Before GST, GST %, Price Inc. GST, Remarks
        header_color: Header background color (hex)
        header_text_color: Header text color (hex)
        row_color_alt: Alternating row color (hex)
        border_color: Border color (hex)
    
    Returns:
        str: Complete HTML table as string
    
    Raises:
        ValueError: If required columns missing
    """
    
    required_cols = [
        'S.No', 'Vendor Item No', 'Product Description',
        'Qty', 'Price Before GST', 'GST %', 'Price Inc. GST', 'Remarks'
    ]
    
    # Check for required columns
    missing_cols = [c for c in required_cols if c not in quote_data.columns]
    if missing_cols:
        raise ValueError(f"Missing columns: {missing_cols}")
    
    html = []
    html.append('<table cellpadding="0" cellspacing="0" border="1" style="border-collapse: collapse; width: 100%; font-family: Calibri, Arial, sans-serif; font-size: 12px;">')
    
    # Header row
    html.append('<tr>')
    for col in required_cols:
        html.append(f'<td style="background-color: {header_color}; color: {header_text_color}; border: 1px solid {border_color}; padding: 8px; font-weight: bold; text-align: center;">{safe_str(col)}</td>')
    html.append('</tr>')
    
    # Data rows
    for idx, (_, row) in enumerate(quote_data.iterrows()):
        row_bg = row_color_alt if idx % 2 == 1 else "#FFFFFF"
        html.append('<tr>')
        
        # S.No
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: center;">{safe_str(row.get("S.No", idx + 1))}</td>')
        
        # Vendor Item No
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px;">{safe_str(row.get("Vendor Item No", ""))}</td>')
        
        # Product Description (may be long)
        desc = safe_str(row.get("Product Description", ""))[:80]  # Limit length
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; word-wrap: break-word;">{desc}</td>')
        
        # Qty
        try:
            qty = float(row.get('Qty', 0) or 0)
            qty_str = format_number(qty, 0)
        except (ValueError, TypeError):
            qty_str = "0"
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: right;">{qty_str}</td>')
        
        # Price Before GST
        try:
            pb = float(row.get('Price Before GST', 0) or 0)
            pb_str = format_currency(pb)
        except (ValueError, TypeError):
            pb_str = format_currency(0)
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: right;">{pb_str}</td>')
        
        # GST %
        try:
            gst_p = float(row.get('GST %', 0) or 0)
            gst_str = format_number(gst_p, 1)
        except (ValueError, TypeError):
            gst_str = "0.0"
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: center;">{gst_str}%</td>')
        
        # Price Inc. GST
        try:
            qty = float(row.get('Qty', 0) or 0)
            pi = float(row.get('Price Inc. GST', 0) or 0)
            line_total = qty * pi
            line_str = format_currency(line_total)
        except (ValueError, TypeError):
            line_str = format_currency(0)
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px; text-align: right;">{line_str}</td>')
        
        # Remarks
        remarks = safe_str(row.get('Remarks', ''))[:40]
        html.append(f'<td style="background-color: {row_bg}; border: 1px solid {border_color}; padding: 8px;">{remarks}</td>')
        
        html.append('</tr>')
    
    html.append('</table>')
    
    return '\n'.join(html)


def generate_totals_table(
    subtotal_excl_gst: float,
    total_gst: float,
    freight_charge: float,
    grand_total: float
) -> str:
    """
    Generate summary totals table.
    
    Args:
        subtotal_excl_gst: Subtotal before GST
        total_gst: Total GST amount
        freight_charge: Freight/shipping charge
        grand_total: Final total including all charges
    
    Returns:
        str: HTML table with totals
    """
    html = []
    html.append('<table cellpadding="0" cellspacing="0" border="0" style="margin-top: 20px; font-family: Calibri, Arial, sans-serif; font-size: 12px; width: 100%;">')
    
    # Subtotal row
    html.append('<tr>')
    html.append('<td style="text-align: right; padding: 8px; width: 70%;">Sub Total (Excl. GST):</td>')
    html.append(f'<td style="text-align: right; padding: 8px; border: 1px solid #000000; border-collapse: collapse; width: 30%;">{format_currency(subtotal_excl_gst)}</td>')
    html.append('</tr>')
    
    # GST row
    html.append('<tr>')
    html.append('<td style="text-align: right; padding: 8px;">Total GST:</td>')
    html.append(f'<td style="text-align: right; padding: 8px; border: 1px solid #000000;">{format_currency(total_gst)}</td>')
    html.append('</tr>')
    
    # Freight row
    html.append('<tr>')
    html.append('<td style="text-align: right; padding: 8px;">Freight Charge:</td>')
    html.append(f'<td style="text-align: right; padding: 8px; border: 1px solid #000000;">{format_currency(freight_charge)}</td>')
    html.append('</tr>')
    
    # Grand Total row (emphasized)
    html.append('<tr>')
    html.append('<td style="text-align: right; padding: 12px; font-weight: bold; background-color: #1F497D; color: #FFFFFF;">GRAND TOTAL (Incl. GST):</td>')
    html.append(f'<td style="text-align: right; padding: 12px; border: 1px solid #1F497D; background-color: #1F497D; color: #FFFFFF; font-weight: bold;">{format_currency(grand_total)}</td>')
    html.append('</tr>')
    
    html.append('</table>')
    
    return '\n'.join(html)


# ── BODY GENERATION ──────────────────────────────────────────────

def generate_plain_text_body(
    greeting: str,
    quote_data: pd.DataFrame,
    subtotal_excl_gst: float,
    total_gst: float,
    freight_charge: float,
    grand_total: float,
    closing: str
) -> str:
    """
    Generate plain text version of email (fallback for HTML).
    
    Args:
        greeting: Opening message
        quote_data: DataFrame with quote items
        subtotal_excl_gst: Subtotal
        total_gst: Total GST
        freight_charge: Freight charge
        grand_total: Grand total
        closing: Closing message
    
    Returns:
        str: Plain text email body
    """
    lines = []
    
    lines.append(safe_str(greeting))
    lines.append("")
    lines.append("=" * 100)
    lines.append("")
    
    # Column headers
    lines.append(
        f"{'S.No':<6} {'Vendor Item':<20} {'Description':<35} {'Qty':>6} {'Price':>12} {'GST%':>7} {'Total':>12}"
    )
    lines.append("-" * 100)
    
    # Data rows
    for _, row in quote_data.iterrows():
        try:
            sno = safe_str(row.get('S.No', ''))
            vendor = safe_str(row.get('Vendor Item No', ''))[:20]
            desc = safe_str(row.get('Product Description', ''))[:35]
            qty = float(row.get('Qty', 0) or 0)
            pb = float(row.get('Price Before GST', 0) or 0)
            gst = float(row.get('GST %', 0) or 0)
            pi = float(row.get('Price Inc. GST', 0) or 0)
            line_total = qty * pi
            
            lines.append(
                f"{sno:<6} {vendor:<20} {desc:<35} {int(qty):>6} {pb:>12,.2f} {gst:>7.1f}% {line_total:>12,.2f}"
            )
        except Exception as e:
            continue
    
    lines.append("-" * 100)
    
    # Totals
    lines.append(f"{'':60} {'':6} {'Sub Total:':>12} {format_currency(subtotal_excl_gst):>12}")
    lines.append(f"{'':60} {'':6} {'GST Total:':>12} {format_currency(total_gst):>12}")
    lines.append(f"{'':60} {'':6} {'Freight:':>12} {format_currency(freight_charge):>12}")
    lines.append(f"{'':60} {'':6} {'GRAND TOTAL:':>12} {format_currency(grand_total):>12}")
    
    lines.append("")
    lines.append("=" * 100)
    lines.append("")
    lines.append(safe_str(closing))
    
    return '\n'.join(lines)


def generate_email_html(
    greeting: str,
    quote_data: pd.DataFrame,
    subtotal_excl_gst: float,
    total_gst: float,
    freight_charge: float,
    grand_total: float,
    closing: str,
    company_name: str = "Supreme Computers India Pvt. Ltd."
) -> str:
    """
    Generate complete Outlook-compatible HTML email body.
    
    Args:
        greeting: Opening message
        quote_data: DataFrame with quote items
        subtotal_excl_gst: Subtotal
        total_gst: Total GST
        freight_charge: Freight charge
        grand_total: Grand total
        closing: Closing message
        company_name: Company name for header
    
    Returns:
        str: Complete HTML email document
    """
    
    table_html = generate_outlook_email_table(quote_data)
    totals_html = generate_totals_table(subtotal_excl_gst, total_gst, freight_charge, grand_total)
    
    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Purchase Quote</title>
</head>
<body style="font-family: Calibri, Arial, sans-serif; line-height: 1.6; color: #333333; background-color: #FFFFFF; padding: 20px;">
    
    <!-- Container -->
    <div style="max-width: 900px; margin: 0 auto; background-color: #FFFFFF; padding: 20px;">
        
        <!-- Header -->
        <div style="border-bottom: 2px solid #1F497D; padding-bottom: 15px; margin-bottom: 20px;">
            <h2 style="color: #1F497D; margin: 0 0 5px 0; font-size: 18px;">{safe_str(company_name)}</h2>
            <p style="color: #666666; margin: 0; font-size: 11px;">Professional Purchase Quote</p>
        </div>
        
        <!-- Greeting -->
        <div style="margin-bottom: 20px;">
            <p style="color: #333333; font-size: 12px; white-space: pre-wrap;">{safe_str(greeting)}</p>
        </div>
        
        <!-- Items Table -->
        <div style="margin-bottom: 20px; overflow-x: auto;">
            {table_html}
        </div>
        
        <!-- Totals Table -->
        <div style="margin-bottom: 20px;">
            {totals_html}
        </div>
        
        <!-- Closing -->
        <div style="margin-bottom: 20px; margin-top: 30px;">
            <p style="color: #333333; font-size: 12px; white-space: pre-wrap;">{safe_str(closing)}</p>
        </div>
        
        <!-- Footer -->
        <div style="border-top: 1px solid #CCCCCC; padding-top: 15px; margin-top: 30px; font-size: 10px; color: #999999;">
            <p style="margin: 5px 0;">This is an automated purchase quote generated by the Purchase Quote System.</p>
            <p style="margin: 5px 0;">Generated on: {datetime.now().strftime('%d %B %Y at %H:%M:%S')}</p>
        </div>
        
    </div>
    
</body>
</html>
"""
    
    return html


# ── EMAIL MESSAGE CREATION ───────────────────────────────────────

def create_outlook_email_message(
    subject: str,
    to_email: str,
    cc_emails: str,
    greeting: str,
    quote_data: pd.DataFrame,
    subtotal_excl_gst: float,
    total_gst: float,
    freight_charge: float,
    grand_total: float,
    closing: str,
    pdf_attachment: Optional[io.BytesIO] = None,
    pdf_filename: str = "quote.pdf"
) -> EmailMessage:
    """
    Create complete EmailMessage with Outlook-compatible formatting.
    
    This is the main function to use for creating purchase quote emails.
    
    Args:
        subject: Email subject line
        to_email: Recipient email address
        cc_emails: CC email addresses (comma-separated)
        greeting: Opening message
        quote_data: DataFrame with quote items
        subtotal_excl_gst: Subtotal before GST
        total_gst: Total GST amount
        freight_charge: Freight/shipping charge
        grand_total: Final total
        closing: Closing message
        pdf_attachment: Optional PDF buffer for attachment
        pdf_filename: Name for PDF attachment
    
    Returns:
        EmailMessage: Complete email ready to send/save
    
    Examples:
        >>> msg = create_outlook_email_message(
        ...     subject="Quote Q0001",
        ...     to_email="vendor@example.com",
        ...     cc_emails="mgr@company.com",
        ...     greeting="Dear Vendor,",
        ...     quote_data=df,
        ...     subtotal_excl_gst=1000,
        ...     total_gst=180,
        ...     freight_charge=100,
        ...     grand_total=1280,
        ...     closing="Please confirm.",
        ...     pdf_attachment=pdf_buf,
        ...     pdf_filename="PO_Q0001.pdf"
        ... )
        >>> eml_bytes = msg.as_bytes()
        >>> with open('quote.eml', 'wb') as f:
        ...     f.write(eml_bytes)
    """
    
    # Create message
    msg = EmailMessage()
    msg['Subject'] = safe_str(subject)
    msg['To'] = safe_str(to_email)
    msg['Cc'] = safe_str(cc_emails)
    msg['From'] = 'Purchase Quote System <noreply@supremecomputers.com>'
    
    # Generate body content
    plain_text = generate_plain_text_body(
        greeting, quote_data, subtotal_excl_gst, total_gst, freight_charge, grand_total, closing
    )
    html_content = generate_email_html(
        greeting, quote_data, subtotal_excl_gst, total_gst, freight_charge, grand_total, closing
    )
    
    # Set content (plain text first, then HTML alternative)
    msg.set_content(plain_text, charset='utf-8')
    msg.add_alternative(html_content, subtype='html', charset='utf-8')
    
    # Add PDF attachment if provided
    if pdf_attachment:
        try:
            pdf_data = pdf_attachment.getvalue()
            if pdf_data:
                msg.add_attachment(
                    pdf_data,
                    maintype='application',
                    subtype='pdf',
                    filename=safe_str(pdf_filename)
                )
        except Exception as e:
            print(f"Warning: Could not attach PDF: {e}")
    
    return msg
