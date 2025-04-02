import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Replace these with your own details
smtp_server = "smtp.ionos.es"  # Replace with your SMTP server
port = 587  # No coma aquí
sender_email = "pruebas@freire-sanchez-valencia.es"  # Your email
sender_password = "Prueba_123@"  # Use environment variables instead!
receiver_email = "info@freire-sanchez-valencia.es"  # Recipient email
subject = "Test HTML Email"
html_file = "email_template.html"  # Your HTML file

# Read the HTML content from a file
try:
    with open(html_file, "r", encoding="utf-8") as file:
        html_content = file.read()
except FileNotFoundError:
    print("Error: El archivo HTML no se encontró.")
    exit()

# Create email message
message = MIMEMultipart("alternative")
message["Subject"] = "Your emaill "
message["From"] = sender_email
message["To"] = receiver_email

# Create a MIMEtext object for the HTML content
html_part = MIMEText(html_content, "html")

# Attach HTML content
message.attach(html_part)

server = None  # Asegurar que server está definido

try:
    # Connect to SMTP server
    server = smtplib.SMTP(smtp_server, port)
    server.ehlo()
    server.starttls()  # Secure connection
    server.ehlo()
    server.login(sender_email, sender_password)
    
    # Send email
    server.sendmail(sender_email, receiver_email, message.as_string())
    print("Email sent successfully!")

except Exception as e:
    print(f"Error: {e}")

finally:
    if server:
        server.quit()

