import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import ssl

st.title("📧 Bulk Email Sender using Excel (Personalized)")

# --- Step 1: Upload Excel File ---
st.header("Step 1: Upload Recipient List")
uploaded_file = st.file_uploader("Upload Excel File with Name, Email (and optional Company) columns", type=["xlsx"])

recipients = []
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    required_cols = {"Name", "Email"}
    if not required_cols.issubset(df.columns):
        st.error("❌ Excel file must have at least 'Name' and 'Email' columns.")
    else:
        # Handle optional company column
        if "Company" not in df.columns:
            df["Company"] = ""  # Add empty if not provided

        recipients = df[["Name", "Email", "Company"]].dropna(subset=["Email"]).to_dict("records")
        st.success(f"✅ Loaded {len(recipients)} recipient(s).")
        st.write(df)

# --- Step 2: Show Email Form Only If File Uploaded ---
if recipients:
    st.header("Step 2: Compose and Send Emails")

    # Email account details
    sender_email = st.text_input("Your Email Address")
    password = st.text_input("Your Email Password / App Password", type="password")

    # Email Content
    subject = st.text_input("Email Subject")
    message_template = st.text_area("Email Message (Use {name} and {company} for personalization)")
    uploaded_image = st.file_uploader("Optional: Upload Image", type=["png", "jpg", "jpeg"])

    if st.button("🚀 Send Emails"):
        if not sender_email or not password:
            st.error("⚠️ Please enter your email and password.")
        elif not subject or not message_template:
            st.error("⚠️ Please enter subject and message.")
        else:
            try:
                context = ssl.create_default_context()
                with smtplib.SMTP("smtp.gmail.com", 587) as server:
                    server.starttls(context=context)
                    server.login(sender_email, password)

                    progress = st.progress(0)
                    status_text = st.empty()

                    for i, rec in enumerate(recipients):
                        recipient_name = rec.get("Name", "")
                        recipient_email = rec.get("Email", "")
                        recipient_company = rec.get("Company", "")

                        # Personalize message safely
                        try:
                            personalized_message = message_template.format(
                                name=recipient_name,
                                company=recipient_company
                            )
                        except KeyError:
                            st.error("⚠️ Error: Please ensure you only use {name} and {company} placeholders.")
                            break

                        msg = MIMEMultipart()
                        msg["From"] = sender_email
                        msg["To"] = recipient_email
                        msg["Subject"] = subject

                        # Attach personalized message
                        msg.attach(MIMEText(f"Hello {recipient_name},\n\n{personalized_message}", "plain"))

                        # Attach image if uploaded
                        if uploaded_image is not None:
                            uploaded_image.seek(0)
                            img = MIMEImage(uploaded_image.read())
                            img.add_header("Content-Disposition", "attachment", filename=uploaded_image.name)
                            msg.attach(img)

                        try:
                            server.sendmail(sender_email, recipient_email, msg.as_string())
                            status_text.text(f"📨 Sent to {recipient_name} ({recipient_email})")
                        except Exception as send_err:
                            status_text.text(f"❌ Failed to {recipient_email}: {send_err}")

                        progress.progress((i + 1) / len(recipients))

                st.success(f"🎉 Personalized emails sent successfully to {len(recipients)} recipients!")
            except Exception as e:
                st.error(f"⚠️ Error: {e}")
