import os
import base64
import json
from openai import OpenAI
from dotenv import load_dotenv

class OpenAIExtractor:
    """Class to handle document analysis using OpenAI's API"""
    
    def __init__(self):
        """Initialize the OpenAI extractor with API credentials"""
        load_dotenv()  # Load environment variables from .env file
        
        # Get OpenAI API key from environment variables
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            print("Warning: OpenAI API key not found in environment variables.")
            print("Please set the OPENAI_API_KEY environment variable.")
        
        self.client = OpenAI(api_key=api_key)
    
    def extract_from_image(self, image_path):
        """
        Extract information from an image using OpenAI's Vision API.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Dictionary with extracted structured data
        """
        if not os.path.exists(image_path):
            print(f"Error: Image file not found at {image_path}")
            return None
        
        # Read the image file and encode as base64
        with open(image_path, "rb") as image_file:
            base64_image = base64.b64encode(image_file.read()).decode('utf-8')
        
        try:
            # Call OpenAI API to analyze the image
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "system",
                        "content": "You are a delivery note parser that extracts structured information from delivery notes. Extract the following information: shipping address, date, sender/from, and list of products with quantities, descriptions, and prices. Format as JSON."
                    },
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": "Parse this delivery note image and extract the structured information. Return ONLY a JSON object with these fields: shipping_address, date, sender, products (array of objects with description, quantity, price), and total_price."},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/png;base64,{base64_image}"
                                }
                            }
                        ]
                    }
                ],
                max_tokens=1500
            )
            
            # Extract the content from the API response
            content = response.choices[0].message.content
            
            # Parse JSON from the response
            # Find JSON object boundaries if there's surrounding text
            json_start = content.find('{')
            json_end = content.rfind('}') + 1
            
            if json_start >= 0 and json_end > json_start:
                json_str = content[json_start:json_end]
                try:
                    result = json.loads(json_str)
                    return result
                except json.JSONDecodeError as e:
                    print(f"Error parsing JSON response: {e}")
                    print(f"Response content: {content}")
                    return self._fallback_parsing(content)
            else:
                print("No JSON found in the response")
                return self._fallback_parsing(content)
                
        except Exception as e:
            if "insufficient_quota" in str(e):
                print("OpenAI API quota exceeded. Please check your billing details or upgrade your plan.")
                print("The application will fall back to OCR-based extraction.")
            elif "model_not_found" in str(e):
                print("The specified OpenAI model is not available. The model may have been deprecated or renamed.")
                print("Please update the model name in the code or contact OpenAI for more information.")
                print("The application will fall back to OCR-based extraction.")
            else:
                print(f"Error calling OpenAI API: {e}")
            return None
    
    def extract_from_pdf(self, pdf_path):
        """
        Extract information from a PDF using OpenAI API.
        For PDF files, we'll convert to images first and then process.
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            Dictionary with extracted structured data
        """
        try:
            from pdf2image import convert_from_path
            import tempfile
            
            # Convert PDF to images
            images = convert_from_path(pdf_path)
            
            # Process first page only for now
            # In a production system, you might want to process all pages
            if images:
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                    temp_path = temp_file.name
                    images[0].save(temp_path, 'PNG')
                    
                # Extract information from the image
                result = self.extract_from_image(temp_path)
                
                # Clean up the temporary file
                os.unlink(temp_path)
                
                return result
            else:
                print(f"No images extracted from PDF: {pdf_path}")
                return None
                
        except Exception as e:
            if "insufficient_quota" in str(e):
                print("OpenAI API quota exceeded. Please check your billing details or upgrade your plan.")
                print("The application will fall back to OCR-based extraction.")
            elif "model_not_found" in str(e):
                print("The specified OpenAI model is not available. The model may have been deprecated or renamed.")
                print("Please update the model name in the code or contact OpenAI for more information.")
                print("The application will fall back to OCR-based extraction.")
            else:
                print(f"Error processing PDF: {e}")
            return None
    
    def _fallback_parsing(self, text):
        """
        Fallback method to parse text when JSON parsing fails.
        Attempts to extract structured data using OpenAI without expecting JSON.
        
        Args:
            text: The text content from the API response
            
        Returns:
            Dictionary with extracted data
        """
        try:
            # Make another API call to parse the text
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {
                        "role": "system",
                        "content": "You are a delivery note parser. Extract structured information and return as JSON."
                    },
                    {
                        "role": "user",
                        "content": f"Parse this delivery note text and return a JSON object with these fields: shipping_address, date, sender, products (array of objects with description, quantity, price), and total_price. Here's the text:\n\n{text}"
                    }
                ],
                response_format={"type": "json_object"}
            )
            
            content = response.choices[0].message.content
            
            try:
                return json.loads(content)
            except json.JSONDecodeError:
                print("Fallback parsing also failed to produce valid JSON")
                return {
                    "shipping_address": "Unknown",
                    "date": "Unknown",
                    "sender": "Unknown",
                    "products": [],
                    "total_price": "0.00"
                }
                
        except Exception as e:
            print(f"Error in fallback parsing: {e}")
            return {
                "shipping_address": "Unknown",
                "date": "Unknown",
                "sender": "Unknown",
                "products": [],
                "total_price": "0.00"
            } 