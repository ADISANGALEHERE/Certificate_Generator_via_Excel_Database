from flask import Flask, render_template, request, send_from_directory
from PIL import Image, ImageDraw, ImageFont
import io
from flask_mail import Mail, Message
import os
from dotenv import load_dotenv
import uuid
import pandas as pd
import datetime

# Load environment variables
load_dotenv()

# Initialize Flask app
app = Flask(__name__)

# Configure Flask-Mail
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USE_SSL'] = False
app.config['MAIL_USERNAME'] = os.getenv('MAIL_USERNAME')
app.config['MAIL_PASSWORD'] = os.getenv('MAIL_PASSWORD')

mail = Mail(app)

# Ensure the uploads directory exists
os.makedirs('uploads', exist_ok=True)

# Route for the index page
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle certificate generation
@app.route('/generate_certificate', methods=['POST'])
def generate_certificate():
    uploaded_file = request.files['excel_file']
    if not uploaded_file:
        return "No file uploaded", 400

    # Generate a unique filename to avoid conflicts
    base_name, ext = os.path.splitext(uploaded_file.filename)
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    unique_excel_file_name = f"{base_name}_{timestamp}{ext}"
    client_excel_file_path = os.path.join('uploads', unique_excel_file_name)

    # Save the uploaded file
    uploaded_file.save(client_excel_file_path)

    # Read the Excel file
    try:
        df = pd.read_excel(client_excel_file_path)
    except Exception as e:
        return f"Error reading Excel file: {str(e)}", 400

    # Check for required columns
    required_columns = ["Name", "Department", "Event", "Event Date", "Email"]
    if not all(col in df.columns for col in required_columns):
        return f"Excel file is missing one or more required columns: {', '.join(required_columns)}", 400

    # Add 'Unique ID' column if not already present
    if 'Unique ID' not in df.columns:
        df.insert(0, 'Unique ID', '')

    # Process each row
    for idx, row in df.iterrows():
        try:
            # Extract participant details
            name = row['Name']
            department = row['Department']
            event = row['Event']
            event_date = row['Event Date']
            recipient_email = row['Email']

            # Generate a unique ID
            unique_id = str(uuid.uuid4())[:8]

            # Load the certificate template
            template_path = 'static/certificate.png'
            img = Image.open(template_path)
            draw = ImageDraw.Draw(img)

            # Load fonts
            font_path = 'path/to/arial.ttf'  # Update this to the correct font path
            font = ImageFont.truetype(font_path, 40)
            small_font = ImageFont.truetype(font_path, 20)

            # Define positions for text on the certificate
            name_position = (151, 823)
            department_position = (1061, 848)
            event_position = (1100, 929)
            date_position = (1435, 977)
            id_position = (img.width - 270, img.height - 1000)

            # Draw text on the certificate
            event_date_str = event_date.strftime("%d-%m-%Y") if isinstance(event_date, datetime.date) else str(event_date)
            draw.text(name_position, name, font=font, fill="black")
            draw.text(department_position, department, font=font, fill="black")
            draw.text(event_position, event, font=font, fill="black")
            draw.text(date_position, event_date_str, font=font, fill="black")
            draw.text(id_position, f"ID: {unique_id}", font=small_font, fill="gray")

            # Save the certificate to a BytesIO object
            output = io.BytesIO()
            img.save(output, format="PNG")
            output.seek(0)

            # Update the 'Unique ID' column in the DataFrame
            df.at[idx, 'Unique ID'] = unique_id

            # Send the certificate via email
            msg = Message(
                "Your Certificate",
                sender=os.getenv('MAIL_USERNAME'),
                recipients=[recipient_email]
            )
            msg.body = (f"Hello {name},\n\n"
                        f"Congratulations on participating in {event} from {department} on {event_date_str}! "
                        f"Please find your certificate attached. Your certificate ID is {unique_id}.")
            msg.attach(f"{unique_id}_certificate.png", "image/png", output.getvalue())
            mail.send(msg)

        except Exception as e:
            return f"Error processing row {idx + 1}: {str(e)}", 500

    # Save the updated Excel file with Unique IDs
    updated_excel_file_name = f"{base_name}_updated_{timestamp}{ext}"
    updated_excel_file_path = os.path.join('uploads', updated_excel_file_name)
    try:
        df.to_excel(updated_excel_file_path, index=False)
    except Exception as e:
        return f"Error saving updated Excel file: {str(e)}", 500

    # Provide a link to download the updated Excel file
    return render_template('download.html', filename=updated_excel_file_name)

# Route to handle downloading the updated Excel file
@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory('uploads', filename, as_attachment=True)

# Run the Flask app
if __name__ == '__main__':
    app.run(debug=True)
