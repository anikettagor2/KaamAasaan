from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import json
import os
from datetime import datetime
import uuid
import io
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image, ImageDraw, ImageFont
import base64
import requests
from fastapi import FastAPI
app = FastAPI()


app = Flask(__name__)
CORS(app)

# Configuration
HANDWRITTEN_REQUESTS_FILE = 'handwritten_requests.json'

# Font URLs (Google Fonts)
FONT_URLS = {
    'ArchitectsDaughter-Regular': 'https://github.com/google/fonts/raw/main/ofl/architectsdaughter/ArchitectsDaughter-Regular.ttf',
    'HomemadeApple-Regular': 'https://github.com/google/fonts/raw/main/ofl/homemadeapple/HomemadeApple-Regular.ttf',
    'Caveat-Regular': 'https://github.com/google/fonts/raw/main/ofl/caveat/Caveat-Regular.ttf',
    'IndieFlower-Regular': 'https://github.com/google/fonts/raw/main/ofl/indieflower/IndieFlower-Regular.ttf',
    'Kalam-Regular': 'https://github.com/google/fonts/raw/main/ofl/kalam/Kalam-Regular.ttf',
    'PatrickHand-Regular': 'https://github.com/google/fonts/raw/main/ofl/patrickhand/PatrickHand-Regular.ttf',
    'PermanentMarker-Regular': 'https://github.com/google/fonts/raw/main/ofl/permanentmarker/PermanentMarker-Regular.ttf',
    'ShadowsIntoLight-Regular': 'https://github.com/google/fonts/raw/main/ofl/shadowsintolight/ShadowsIntoLight-Regular.ttf'
}

def download_and_register_font(font_name):
    """Download and register a font for use in PDFs"""
    try:
        # Check if font is already registered
        if font_name in pdfmetrics.getRegisteredFontNames():
            return True
        
        # Download font if not already downloaded
        font_path = f'fonts/{font_name}.ttf'
        if not os.path.exists(font_path):
            os.makedirs('fonts', exist_ok=True)
            
            print(f"Downloading font: {font_name}")
            # Download font from Google Fonts
            response = requests.get(FONT_URLS[font_name], timeout=30)
            if response.status_code == 200:
                with open(font_path, 'wb') as f:
                    f.write(response.content)
                print(f"Font downloaded successfully: {font_name}")
            else:
                print(f"Failed to download font {font_name}: {response.status_code}")
                return False
        
        # Register font with reportlab
        pdfmetrics.registerFont(TTFont(font_name, font_path))
        print(f"Font registered successfully: {font_name}")
        return True
        
    except Exception as e:
        print(f"Error downloading/registering font {font_name}: {e}")
        return False

def load_requests():
    """Load existing requests from file"""
    if os.path.exists(HANDWRITTEN_REQUESTS_FILE):
        with open(HANDWRITTEN_REQUESTS_FILE, 'r') as f:
            return json.load(f)
    return []

def save_requests(requests):
    """Save requests to file"""
    with open(HANDWRITTEN_REQUESTS_FILE, 'w') as f:
        json.dump(requests, f, indent=2)

@app.route('/api/submit-handwritten-request', methods=['POST'])
def submit_handwritten_request():
    """Handle handwritten assignment generation requests"""
    try:
        data = request.json
        
        # Validate required fields
        required_fields = ['projectTitle', 'assignmentContent', 'studentName', 'rollNumber', 'subject', 'professor']
        for field in required_fields:
            if not data.get(field):
                return jsonify({'success': False, 'error': f'Missing required field: {field}'}), 400
        
        # Generate a unique ID for this request
        request_id = str(uuid.uuid4())
        
        # Store the request data
        request_data = {
            'id': request_id,
            'projectTitle': data['projectTitle'],
            'assignmentContent': data['assignmentContent'],
            'studentName': data['studentName'],
            'rollNumber': data['rollNumber'],
            'subject': data['subject'],
            'professor': data['professor'],
            'handwritingFont': data.get('handwritingFont', 'Homemade Apple'),
            'fontSize': data.get('fontSize', 12),
            'inkColor': data.get('inkColor', '#000000'),
            'paperLines': data.get('paperLines', True),
            'paperMargins': data.get('paperMargins', True),
            'created_at': datetime.now().isoformat(),
            'status': 'completed'
        }
        
        # Save to file
        requests = load_requests()
        requests.append(request_data)
        save_requests(requests)
        
        return jsonify({
            'success': True,
            'message': 'Handwritten assignment generated successfully!',
            'request_id': request_id,
            'assignment_data': request_data
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/generate-handwritten/<request_id>/pdf', methods=['GET'])
def generate_handwritten_pdf(request_id):
    """Generate and download handwritten assignment as PDF"""
    try:
        # Load the stored request data
        requests = load_requests()
        request_data = None
        
        for req in requests:
            if req['id'] == request_id:
                request_data = req
                break
        
        if not request_data:
            return jsonify({'success': False, 'error': 'Request not found'}), 404
        
        # Use the stored data
        project_title = request_data['projectTitle']
        assignment_content = request_data['assignmentContent']
        student_name = request_data['studentName']
        roll_number = request_data['rollNumber']
        subject = request_data['subject']
        professor = request_data['professor']
        handwriting_font = request_data['handwritingFont']
        font_size = request_data['fontSize']
        ink_color = request_data['inkColor']
        paper_lines = request_data['paperLines']
        paper_margins = request_data['paperMargins']
        
        # Create PDF with handwritten-style content
        return create_handwritten_pdf(
            project_title, assignment_content, student_name, roll_number, 
            subject, professor, handwriting_font, font_size, ink_color, 
            paper_lines, paper_margins
        )
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

def create_handwritten_pdf(project_title, assignment_content, student_name, roll_number, 
                          subject, professor, handwriting_font, font_size, ink_color, 
                          paper_lines, paper_margins):
    """Create a PDF that looks like handwritten content with lined paper"""
    try:
        # Create a buffer to store the PDF
        buffer = io.BytesIO()
        
        # Create the PDF object with letter size
        doc = SimpleDocTemplate(buffer, pagesize=letter)
        styles = getSampleStyleSheet()
        
        # Convert hex color to RGB
        def hex_to_rgb(hex_color):
            hex_color = hex_color.lstrip('#')
            return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        ink_rgb = hex_to_rgb(ink_color)
        
        # Map font names to actual handwritten fonts
        font_mapping = {
            'Architects Daughter': 'ArchitectsDaughter-Regular',
            'Homemade Apple': 'HomemadeApple-Regular',
            'Caveat': 'Caveat-Regular',
            'Indie Flower': 'IndieFlower-Regular',
            'Kalam': 'Kalam-Regular',
            'Patrick Hand': 'PatrickHand-Regular',
            'Permanent Marker': 'PermanentMarker-Regular',
            'Shadows Into Light': 'ShadowsIntoLight-Regular'
        }
        
        # Get the actual font name
        actual_font = font_mapping.get(handwriting_font, 'HomemadeApple-Regular')
        
        print(f"Using font: {actual_font}")
        
        # Download and register the font
        if not download_and_register_font(actual_font):
            print(f"Failed to download font {actual_font}, using fallback")
            # Fallback to a default font if download fails
            actual_font = 'Helvetica'
        
        # Create custom styles that use handwritten fonts
        title_style = ParagraphStyle(
            'HandwrittenTitle',
            parent=styles['Heading1'],
            fontSize=font_size + 4,
            spaceAfter=30,
            alignment=1,  # Center alignment
            fontName=actual_font,
            textColor=ink_rgb
        )
        
        content_style = ParagraphStyle(
            'HandwrittenContent',
            parent=styles['Normal'],
            fontSize=font_size,
            spaceAfter=12,
            leading=font_size * 1.5,  # Line spacing
            fontName=actual_font,
            textColor=ink_rgb
        )
        
        header_style = ParagraphStyle(
            'HandwrittenHeader',
            parent=styles['Normal'],
            fontSize=font_size - 2,
            spaceAfter=6,
            fontName=actual_font,
            textColor=ink_rgb
        )
        
        # Create a custom page template with lined background
        def create_lined_page(canvas, doc):
            canvas.saveState()
            
            # Set background to white
            canvas.setFillColorRGB(1, 1, 1)
            canvas.rect(0, 0, letter[0], letter[1], fill=1)
            
            if paper_lines:
                # Draw horizontal lines with better spacing
                canvas.setStrokeColorRGB(0.9, 0.9, 0.9)  # Very light gray lines
                canvas.setLineWidth(0.3)
                
                # Calculate line spacing based on font size
                line_spacing = font_size * 1.8  # Slightly more spacing for readability
                start_y = letter[1] - 120  # Start below header
                end_y = 60  # Leave margin at bottom
                
                y = start_y
                while y > end_y:
                    canvas.line(60, y, letter[0] - 60, y)
                    y -= line_spacing
            
            if paper_margins:
                # Draw left margin line
                canvas.setStrokeColorRGB(1, 0, 0)  # Red margin line
                canvas.setLineWidth(1.5)
                canvas.line(60, 60, 60, letter[1] - 60)
                
                # Draw top and bottom margins
                canvas.setStrokeColorRGB(0.8, 0.8, 0.8)
                canvas.setLineWidth(0.5)
                canvas.line(60, letter[1] - 60, letter[0] - 60, letter[1] - 60)  # Top
                canvas.line(60, 60, letter[0] - 60, 60)  # Bottom
            
            canvas.restoreState()
        
        # Create page template
        from reportlab.platypus import PageTemplate, Frame
        page_template = PageTemplate(
            id='lined_page',
            frames=[Frame(60, 60, letter[0]-120, letter[1]-120, id='normal')],
            onPage=create_lined_page
        )
        doc.addPageTemplates([page_template])
        
        # Parse content and create PDF elements
        elements = []
        
        # Add title
        elements.append(Paragraph(project_title, title_style))
        elements.append(Spacer(1, 20))
        
        # Add header information
        elements.append(Paragraph(f"Student Name: {student_name}", header_style))
        elements.append(Paragraph(f"Roll Number: {roll_number}", header_style))
        elements.append(Paragraph(f"Subject: {subject}", header_style))
        elements.append(Paragraph(f"Professor: {professor}", header_style))
        elements.append(Paragraph(f"Date: {datetime.now().strftime('%B %d, %Y')}", header_style))
        elements.append(Spacer(1, 30))
        
        # Add assignment content
        # Split content into paragraphs
        paragraphs = assignment_content.split('\n\n')
        for paragraph in paragraphs:
            if paragraph.strip():
                elements.append(Paragraph(paragraph.strip(), content_style))
                elements.append(Spacer(1, 12))
        
        # Build PDF
        doc.build(elements)
        buffer.seek(0)
        
        # Generate filename
        filename = f"{project_title.replace(' ', '_')}_handwritten_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=filename,
            mimetype='application/pdf'
        )
        
    except Exception as e:
        return jsonify({'success': False, 'error': f'PDF generation failed: {str(e)}'}), 500

@app.route('/api/handwritten-requests', methods=['GET'])
def get_handwritten_requests():
    """Get all handwritten assignment requests (for admin dashboard)"""
    try:
        requests = load_requests()
        return jsonify({
            'success': True,
            'requests': requests
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/update-handwritten-status', methods=['POST'])
def update_handwritten_status():
    """Update handwritten request status (for admin)"""
    try:
        data = request.json
        request_id = data.get('id')
        new_status = data.get('status')
        
        if not request_id or not new_status:
            return jsonify({'success': False, 'error': 'Missing id or status'}), 400
        
        requests = load_requests()
        
        for req in requests:
            if req['id'] == request_id:
                req['status'] = new_status
                req['updated_at'] = datetime.now().isoformat()
                break
        
        save_requests(requests)
        
        return jsonify({'success': True, 'message': 'Status updated successfully'})
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001)

