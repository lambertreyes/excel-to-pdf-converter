from http.server import BaseHTTPRequestHandler
import io
import base64
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            # Read all available data
            content_length = self.headers.get('Content-Length')
            if content_length:
                body = self.rfile.read(int(content_length))
            else:
                # Try to read without content length
                body = self.rfile.read()
            
            if not body or len(body) == 0:
                self.send_response(400)
                self.send_header('Content-Type', 'text/plain')
                self.end_headers()
                self.wfile.write(b"Empty body received")
                return
            
            # Load workbook
            wb = load_workbook(io.BytesIO(body), data_only=True)
            sheet = wb.active
            
            # Get all data
            data = []
            for row in sheet.iter_rows(values_only=True):
                data.append([str(c) if c is not None else '' for c in row])
            
            # Create PDF
            pdf_buf = io.BytesIO()
            doc = SimpleDocTemplate(pdf_buf, pagesize=letter)
            
            # Build table
            if data:
                t = Table(data)
                t.setStyle(TableStyle([
                    ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                    ('FONTSIZE', (0,0), (-1,-1), 9),
                ]))
                doc.build([t])
            
            # Return PDF as base64
            pdf_bytes = pdf_buf.getvalue()
            pdf_base64 = base64.b64encode(pdf_bytes).decode()
            
            self.send_response(200)
            self.send_header('Content-Type', 'text/plain')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(pdf_base64.encode())
            
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'text/plain')
            self.end_headers()
            error_msg = f"Error: {str(e)}"
            self.wfile.write(error_msg.encode())
    
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
