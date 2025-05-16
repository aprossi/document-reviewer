#!/usr/bin/env python3
"""
ODT Enhanced Feedback Tool with Side Comments
Provides feedback on OpenDocument Text (.odt) files using LLMs via Ollama's REST API
Creates a duplicate document with comments in the margins
Processes text in paragraphs, tables, and text boxes
"""

import os
import re
import argparse
import sys
import time
import requests
import json
import shutil
import tempfile
import zipfile
from lxml import etree

def check_odfpy_installed():
    """Check if odfpy is installed and return status"""
    try:
        import odf
        from odf.opendocument import load, OpenDocumentText
        from odf.text import P, Span, H, LineBreak
        from odf.style import Style, TextProperties, ParagraphProperties
        from odf.table import Table, TableRow, TableCell
        from odf import teletype
        
        # Don't try to import annotation module which may not exist in older versions
        print(f"✅ odfpy found! Using Python from: {sys.executable}")
        
        # Check if annotation module exists
        try:
            from odf.annotation import Annotation
            print("✅ odf.annotation module found")
            return True
        except ImportError:
            print("⚠️ odf.annotation module not found, will use custom implementation")
            return True  # Still return True, we'll use our custom implementation
            
    except ImportError as e:
        print(f"❌ Import error: {str(e)}")
        print("Please install the required packages:")
        print("pip install odfpy")
        return False

# Check if odfpy is installed
if not check_odfpy_installed():
    sys.exit(1)

# Import the required modules after verification
from odf.opendocument import load, OpenDocumentText
from odf.text import P, Span, H, LineBreak
from odf.style import Style, TextProperties, ParagraphProperties
from odf.table import Table, TableRow, TableCell
from odf import teletype
from odf.element import Element
from odf.namespaces import OFFICENS, TEXTNS, DCNS

# Custom implementation of missing annotation classes if needed
try:
    from odf.annotation import Annotation, Creator, Date, AnnotationBody
    USE_CUSTOM_ANNOTATION = False
except ImportError:
    USE_CUSTOM_ANNOTATION = True
    
    # Custom implementations of annotation classes
    class Annotation(Element):
        def __init__(self):
            Element.__init__(self, qname=(OFFICENS, "annotation"))
            
    class Creator(Element):
        def __init__(self):
            Element.__init__(self, qname=(DCNS, "creator"))
            
    class Date(Element):
        def __init__(self):
            Element.__init__(self, qname=(DCNS, "date"))
            
    # Note: We don't use AnnotationBody in our code anymore
    # Instead we use regular P elements inside annotations

def main():
    """Main function to run the script"""
    parser = argparse.ArgumentParser(
        description="Add side comments to an ODT document using LLMs with Ollama",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("odt_path", help="Path to the ODT file to analyze", nargs="?")
    parser.add_argument("--model", "-m", help="Specific model to use (e.g., llama3.1, mistral, phi3)", default=None)
    parser.add_argument("--verbose", "-v", action="store_true", help="Print detailed progress")
    parser.add_argument("--list-models", "-l", action="store_true", help="List available models and exit")
    parser.add_argument("--system-prompt", "-s", help="Path to a custom system prompt file")
    parser.add_argument("--creative", "-c", action="store_true", 
                       help="Use a higher temperature for more creative feedback")
    parser.add_argument("--api-host", help="Ollama API host (e.g., http://localhost:11434)", 
                       default="http://localhost:11434")
    parser.add_argument("--no-summary", action="store_true", 
                       help="Don't create a separate summary document")
    parser.add_argument("--version", action="version", version="ODT Comments Feedback Tool v1.0.0")
    args = parser.parse_args()
    
    # Handle the list-models argument
    if args.list_models:
        list_available_models(api_host=args.api_host)
        return
    
    if not args.odt_path:
        if not args.list_models:
            parser.print_help()
        return
        
    if not os.path.exists(args.odt_path):
        print(f"❌ Error: File not found - {args.odt_path}")
        return
    
    if not args.odt_path.lower().endswith(".odt"):
        print(f"❌ Error: File must be an .odt file - {args.odt_path}")
        return
    
    # Check Ollama setup with the specified model if provided
    model_name = confirm_ollama_setup(args.model, api_host=args.api_host)
    if not model_name:
        return
    
    # Check if custom system prompt is provided
    system_prompt = None
    if args.system_prompt:
        if os.path.exists(args.system_prompt):
            with open(args.system_prompt, 'r') as f:
                system_prompt = f.read()
            print(f"Using custom system prompt from: {args.system_prompt}")
        else:
            print(f"⚠️ Warning: System prompt file not found: {args.system_prompt}")
            print("Using default system prompt instead.")
    
    # Analyze the document with side comments
    analyze_odt_document(
        odt_path=args.odt_path, 
        model_name=model_name,
        verbose=args.verbose,
        temperature=0.7 if args.creative else 0.2,
        api_host=args.api_host,
        system_prompt=system_prompt,
        summary_doc=not args.no_summary
    )

def extract_text_from_odt(odt_path):
    """
    Extract text with location information from ODT document
    Args:
        odt_path: Path to the ODT file
    Returns:
        List of dictionaries with text content and detailed location info
    """
    text_elements = []
    
    # Load the ODT document
    print(f"Loading ODT document: {odt_path}")
    doc = load(odt_path)
    
    # Process main document paragraphs
    print("Extracting text from main document...")
    
    # Get all paragraph elements
    paragraphs = doc.getElementsByType(P)
    for i, para in enumerate(paragraphs):
        # Extract text from paragraph
        text = teletype.extractText(para)
        if text.strip():
            text_elements.append({
                'text': text,
                'type': 'paragraph',
                'location': {
                    'paragraph_index': i,
                    'element': para  # Store the element for reference
                },
                'display': f"Paragraph {i+1}"
            })
    
    # Process headings - fix for older odfpy versions
    print("Extracting text from headings...")
    headings = []
    all_headings = doc.getElementsByType(H)
    
    # Manually filter headings by outline level attribute
    for heading_level in range(1, 6):  # H1 to H5
        for h in all_headings:
            # Check if the heading has the right outline level
            if h.getAttribute('outlinelevel') == str(heading_level):
                headings.append(h)
    
    for i, heading in enumerate(headings):
        text = teletype.extractText(heading)
        if text.strip():
            text_elements.append({
                'text': text,
                'type': 'heading',
                'location': {
                    'heading_index': i,
                    'element': heading
                },
                'display': f"Heading {i+1}"
            })
    
    # Process tables
    print("Extracting text from tables...")
    tables = doc.getElementsByType(Table)
    
    for table_idx, table in enumerate(tables):
        rows = table.getElementsByType(TableRow)
        for row_idx, row in enumerate(rows):
            cells = row.getElementsByType(TableCell)
            for cell_idx, cell in enumerate(cells):
                # Get paragraphs within cell
                cell_paragraphs = cell.getElementsByType(P)
                for para_idx, para in enumerate(cell_paragraphs):
                    text = teletype.extractText(para)
                    if text.strip():
                        text_elements.append({
                            'text': text,
                            'type': 'table',
                            'location': {
                                'table_index': table_idx,
                                'row_index': row_idx,
                                'cell_index': cell_idx,
                                'paragraph_index': para_idx,
                                'element': para
                            },
                            'display': f"Table {table_idx+1}, Row {row_idx+1}, Cell {cell_idx+1}, Para {para_idx+1}"
                        })
    
    print(f"Extracted {len(text_elements)} text elements")
    return text_elements, doc

def add_comment_to_odt_element(doc, element, comment_text, author="AI Editor"):
    """
    Add a comment to an element in an ODT document
    Args:
        doc: ODT document object
        element: The element to comment on
        comment_text: The text of the comment
        author: The comment author name
    """
    try:
        # Create annotation elements
        annotation = Annotation()
        
        # Add creator (author)
        creator_elem = Creator()
        creator_elem.addText(author)
        annotation.appendChild(creator_elem)
        
        # Add date
        date_elem = Date()
        date_str = time.strftime("%Y-%m-%dT%H:%M:%S")
        date_elem.addText(date_str)
        annotation.appendChild(date_elem)
        
        # Add annotation body with comment text
        annotation_body = P()  # Use regular paragraph for body in custom implementation
        annotation_body.addText(comment_text)
        annotation.appendChild(annotation_body)
        
        # Add the annotation to the element - safer approach
        try:
            # If it's a paragraph or heading, we can add directly
            if element.qname[1] in ('p', 'h'):
                # Try to insert annotation at the beginning
                try:
                    # Create a simple clone of the paragraph with just the annotation
                    new_para = P()
                    
                    # Add the annotation at the beginning
                    new_para.appendChild(annotation)
                    
                    # Add the original text
                    text = teletype.extractText(element)
                    if text:
                        new_para.addText(text)
                    
                    # Replace the original element with our new one
                    if element.parentNode:
                        element.parentNode.insertBefore(new_para, element)
                        try:
                            element.parentNode.removeChild(element)
                        except:
                            # If removal fails, just continue with both elements
                            pass
                    else:
                        # If no parent, try to add to document
                        doc.text.appendChild(new_para)
                except Exception as elem_err:
                    # Fallback: just append the annotation to the element
                    print(f"Using fallback method for comment: {str(elem_err)}")
                    element.appendChild(annotation)
            else:
                # For other elements, try to find a parent paragraph
                parent = element
                while parent is not None and parent.qname[1] not in ('p', 'h'):
                    parent = parent.parentNode
                
                if parent is not None:
                    parent.appendChild(annotation)
                else:
                    # Create a new paragraph for the annotation
                    p = P()
                    p.appendChild(annotation)
                    
                    # Try to add it to the document
                    if hasattr(doc, 'text'):
                        doc.text.appendChild(p)
                    else:
                        # Last resort - add to document body somehow
                        for body in doc.getElementsByType(doc.getElementsByType(doc.root)[0].__class__):
                            body.appendChild(p)
                            break
            
            return annotation
            
        except Exception as insert_err:
            print(f"Error adding comment: {str(insert_err)}")
            return None
            
    except Exception as e:
        print(f"Error creating comment: {str(e)}")
        return None

def confirm_ollama_setup(model_name=None, api_host=None):
    """
    Check if Ollama is running and the requested model is available
    Args:
        model_name: Optional specific model to check for
        api_host: Optional API host address (e.g. http://localhost:11434)
    Returns:
        model_name if found, or False if not found
    """
    print("Checking Ollama setup...")
    api_host = api_host or "http://localhost:11434"
    
    try:
        # Try direct API call with requests
        print(f"Using direct API call to: {api_host}")
        try:
            response = requests.get(f"{api_host}/api/tags")
            
            if response.status_code == 200:
                data = response.json()
                available_models = []
                
                for model in data.get("models", []):
                    model_name_str = model.get("name", "")
                    if model_name_str:
                        available_models.append(model_name_str)
                
                if not available_models:
                    print("\n⚠️  No models found in Ollama!")
                    print("Please pull a model first, for example:")
                    print("ollama pull llama3.1")
                    return False
                
                # Print available models for debugging
                print(f"Found {len(available_models)} models: {', '.join(available_models[:5])}...")
                
                # If a specific model is requested, check if it's available
                if model_name:
                    if model_name in available_models:
                        print(f"✅ Found requested model: {model_name}")
                        return model_name
                    else:
                        # Check if the model name might be a partial match (without tags)
                        matching_models = [m for m in available_models if m.startswith(f"{model_name}:") or m == model_name]
                        if matching_models:
                            print(f"✅ Found matching model: {matching_models[0]}")
                            return matching_models[0]
                        
                        print(f"\n⚠️  Requested model '{model_name}' not found in Ollama!")
                        print(f"Available models: {', '.join(available_models)}")
                        return False
                
                # If no specific model requested, prefer Llama models in this order:
                preferred_models = [
                    "llama3.1", "llama3.2", "llama3", 
                    "mistral", "gemma3", "phi3", 
                    "deepseek-r1"
                ]
                
                for preferred in preferred_models:
                    matching = [m for m in available_models if m.startswith(f"{preferred}:") or m == preferred]
                    if matching:
                        print(f"✅ Using model: {matching[0]}")
                        return matching[0]
                
                # If no preferred models, just use the first available
                print(f"✅ Using model: {available_models[0]}")
                return available_models[0]
            else:
                print(f"\n❌ API returned status code {response.status_code}")
                print(f"Response: {response.text}")
                return False
        
        except Exception as e:
            print(f"\n❌ Direct API call failed: {str(e)}")
            return False
    
    except Exception as e:
        print(f"\n❌ Error connecting to Ollama: {str(e)}")
        print("Please make sure Ollama is installed and running.")
        print("Installation instructions: https://ollama.com/download")
        print("\nDiagnostic information:")
        print("- Try running 'ollama list' from the command line")
        print("- Check if Ollama is listening on port 11434")
        print("- If running on a different port, use the --api-host option")
        print("- Check Ollama logs for errors")
        return False

def list_available_models(api_host=None):
    """
    List all available models in Ollama and exit
    Args:
        api_host: Optional API host address
    """
    api_host = api_host or "http://localhost:11434"
    
    try:
        # Try direct API call with requests
        print(f"Connecting to Ollama API at: {api_host}")
        
        try:
            response = requests.get(f"{api_host}/api/tags")
            
            if response.status_code == 200:
                data = response.json()
                available_models = data.get("models", [])
                
                # Sort models by name
                available_models.sort(key=lambda x: x.get("name", ""))
                
                # Print in a table format
                print("\n🤖 Available Models in Ollama:")
                print("-" * 80)
                print(f"{'MODEL NAME':<40} {'SIZE':<10} {'MODIFIED':<20}")
                print("-" * 80)
                
                for model in available_models:
                    model_name = model.get("name", "")
                    
                    # Format the size in a human-readable format
                    size_bytes = model.get("size", 0)
                    size_gb = size_bytes / 1_000_000_000
                    size_str = f"{size_gb:.1f} GB"
                    
                    # Format the modified date
                    modified = model.get("modified_at", "")
                    if "T" in modified:
                        modified = modified.split("T")[0]
                    
                    print(f"{model_name:<40} {size_str:<10} {modified:<20}")
                
                print("-" * 80)
                print(f"Total models: {len(available_models)}")
                print(f"\nOllama API host: {api_host}")
                print("\nUse --model MODEL_NAME to select a specific model for analysis.")
                print("Example: python ollama-odt-comments-enhanced.py document.odt --model mistral")
                
            else:
                print(f"\n❌ API returned status code {response.status_code}")
                print(f"Response: {response.text}")
        
        except Exception as e:
            print(f"\n❌ Error connecting to Ollama API: {str(e)}")
            print(f"Tried to connect to: {api_host}")
            print("Please make sure Ollama is installed and running.")
    
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
    
    sys.exit(0)

def create_odt_summary_doc(text_elements, overall_feedback, model_name):
    """
    Create a summary ODT document with feedback
    Args:
        text_elements: List of text elements with feedback
        overall_feedback: Overall document feedback
        model_name: Name of the model used
    Returns:
        OpenDocumentText object
    """
    summary_doc = OpenDocumentText()
    
    # Add title
    title = H(outlinelevel=1)
    title.addText("Document Feedback Summary")
    summary_doc.text.appendChild(title)
    
    # Add info paragraph
    info_para = P()
    info_para.addText(f"Analysis by: {model_name}")
    summary_doc.text.appendChild(info_para)
    
    date_para = P()
    date_para.addText(f"Date: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    summary_doc.text.appendChild(date_para)
    
    # Add overall assessment
    assessment_heading = H(outlinelevel=2)
    assessment_heading.addText("Overall Assessment")
    summary_doc.text.appendChild(assessment_heading)
    
    # Split overall feedback by lines and add them
    for line in overall_feedback.split('\n'):
        if line.strip():
            para = P()
            para.addText(line)
            summary_doc.text.appendChild(para)
    
    # Add content-by-content section
    content_heading = H(outlinelevel=2)
    content_heading.addText("Content-by-Content Feedback")
    summary_doc.text.appendChild(content_heading)
    
    # Add feedback for each element with feedback
    for elem in text_elements:
        if 'feedback' in elem and elem['feedback']:
            # Add source info
            source_para = P()
            source_span = Span(stylename="Bold")
            source_span.addText(f"{elem['type']} {elem['display']}: ")
            source_para.appendChild(source_span)
            
            # Add excerpt
            excerpt = elem['text'][:100] + "..." if len(elem['text']) > 100 else elem['text']
            source_para.addText(excerpt)
            summary_doc.text.appendChild(source_para)
            
            # Add feedback
            feedback_para = P(stylename="Quote")
            feedback_para.addText(f"✎ {elem['feedback']}")
            summary_doc.text.appendChild(feedback_para)
            
            # Add space
            summary_doc.text.appendChild(P())
    
    return summary_doc

def analyze_odt_document(odt_path, model_name="llama3.1", verbose=False, 
                        temperature=0.2, api_host=None, system_prompt=None,
                        summary_doc=True):
    """
    Process an ODT document with LLM and add side comments
    Args:
        odt_path: Path to the ODT file
        model_name: Ollama model name to use
        verbose: Whether to print detailed progress
        temperature: Model temperature (higher = more creative)
        api_host: Optional API host address
        system_prompt: Custom system prompt to use (if None, uses default)
        summary_doc: Whether to create a separate summary document
    Returns:
        Path to the created commented document and summary document (if created)
    """
    api_host = api_host or "http://localhost:11434"
    
    try:
        # Extract text elements and get the document
        text_elements, doc = extract_text_from_odt(odt_path)
        
        # Create duplicate document for comments
        base_name = os.path.basename(odt_path)
        name_without_ext = os.path.splitext(base_name)[0]
        model_short_name = model_name.split(':')[0].replace("/", "-")
        
        # Create paths for the output files
        commented_path = os.path.join(os.path.dirname(odt_path), 
                                    f"{name_without_ext}_comments_{model_short_name}.odt")
        
        # Get text for full document analysis
        all_text = [elem['text'] for elem in text_elements]
        full_text = "\n\n".join(all_text)
        char_count = len(full_text)
        
        # Count by type
        main_paragraphs = sum(1 for elem in text_elements if elem['type'] == 'paragraph')
        headings = sum(1 for elem in text_elements if elem['type'] == 'heading')
        table_cells = sum(1 for elem in text_elements if elem['type'] == 'table')
        
        print(f"Document contains {char_count} characters across {len(text_elements)} text elements")
        print(f"  - Main paragraphs: {main_paragraphs}")
        print(f"  - Headings: {headings}")
        print(f"  - Table cells: {table_cells}")
        
        if char_count == 0:
            print("❌ Document appears to be empty. Please check the file.")
            return None
        
        # Use default system prompt if none provided
        if system_prompt is None:
            system_prompt = '''You are a professional editor providing detailed stylistic feedback.
            Analyze the document for:
            1. Writing style and tone
            2. Clarity and conciseness
            3. Sentence structure and flow
            4. Word choice and vocabulary
            5. Overall impressions and suggestions
            
            Format your feedback as follows:
            
            # Overall Assessment
            [General feedback about the entire document]
            
            # Specific Issues
            1. [First specific issue - be precise]
            2. [Second specific issue]
            [etc.]
            
            # Recommended Improvements
            1. [First recommendation]
            2. [Second recommendation]
            [etc.]
            '''
        
        # Get overall stylistic feedback from the LLM using direct API calls
        print(f"Getting overall document assessment using {model_name}...")
        print(f"Temperature: {temperature:.1f}" + (" (creative mode)" if temperature > 0.5 else ""))
        start_time = time.time()
        
        # Prepare messages
        messages = [
            {
                'role': 'system',
                'content': system_prompt
            },
            {
                'role': 'user',
                'content': f"Please provide stylistic feedback on this document:\n\n{full_text}"
            }
        ]
        
        # Make API call
        try:
            api_data = {
                "model": model_name,
                "messages": messages,
                "options": {
                    "temperature": temperature
                },
                "stream": False
            }
            
            response = requests.post(
                f"{api_host}/api/chat",
                json=api_data
            )
            
            if response.status_code != 200:
                print(f"\n❌ API returned status code {response.status_code}")
                print(f"Response: {response.text}")
                return None
            
            response_data = response.json()
            overall_feedback = response_data.get("message", {}).get("content", "")
            
            if not overall_feedback:
                print("\n❌ No feedback received from model")
                return None
            
        except Exception as e:
            print(f"\n❌ Error calling Ollama API: {str(e)}")
            return None
        
        analysis_time = time.time() - start_time
        print(f"Overall assessment completed in {analysis_time:.1f} seconds")
        
        # Count text elements that will be analyzed
        substantive_text = [elem for elem in text_elements if len(elem['text'].strip()) > 30]
        total_to_analyze = len(substantive_text)
        
        if total_to_analyze > 20:
            print(f"⚠️  Document has {total_to_analyze} substantive text elements.")
            print("   Analysis may take some time.")
            
            if total_to_analyze > 50:
                while True:
                    response = input("Document is very large. Continue? (y/n): ").lower()
                    if response in ('y', 'yes'):
                        break
                    elif response in ('n', 'no'):
                        print("Operation cancelled.")
                        return None
        
        # Process each text element
        analyzed_count = 0
        feedback_count = 0
        comment_errors = 0
        print("Analyzing individual text elements and adding comments...")
        
        # Set to track which elements we've already processed
        processed_elements = set()
        
        # Process each text element
        for i, element in enumerate(text_elements):
            text_content = element['text']
            element_type = element['type']
            display = element['display']
            
            # Skip already processed elements or empty text
            element_key = f"{i}:{text_content[:50]}"
            if element_key in processed_elements or not text_content.strip():
                continue
            
            processed_elements.add(element_key)
            
            # Only analyze substantive text
            if len(text_content.strip()) > 30:
                analyzed_count += 1
                progress = f"[{analyzed_count}/{total_to_analyze}]"
                
                if verbose:
                    print(f"{progress} Analyzing {element_type} {display}...")
                else:
                    # Print progress bar
                    progress_pct = analyzed_count*100//total_to_analyze
                    progress_bar = "█" * (progress_pct//5) + "░" * (20 - progress_pct//5)
                    sys.stdout.write(f"\r{progress} Analyzing content... {progress_bar} {progress_pct}%")
                    sys.stdout.flush()
                
                # Prepare messages for text analysis
                content_messages = [
                    {
                        'role': 'system',
                        'content': '''You are a professional writing coach providing targeted feedback.
                        Review the text and identify any grammar, style, or clarity issues.
                        If you find issues, provide ONE clear, specific suggestion that would most improve it.
                        Be precise and helpful. Limit your feedback to 1-2 sentences.
                        If the text has no significant issues, respond with "No issues found."'''
                    },
                    {
                        'role': 'user',
                        'content': f"Review this text (from {element_type} {display}):\n\n{text_content}"
                    }
                ]
                
                # Get feedback for this text element using direct API
                try:
                    api_data = {
                        "model": model_name,
                        "messages": content_messages,
                        "options": {
                            "temperature": temperature
                        },
                        "stream": False
                    }
                    
                    response = requests.post(
                        f"{api_host}/api/chat",
                        json=api_data
                    )
                    
                    if response.status_code != 200:
                        raise RuntimeError(f"API returned status code {response.status_code}")
                    
                    response_data = response.json()
                    content_feedback = response_data.get("message", {}).get("content", "").strip()
                    
                    # Add feedback if there's something to improve
                    if content_feedback and "No issues found" not in content_feedback:
                        feedback_count += 1
                        
                        # Store the feedback in the element for summary document
                        element['feedback'] = content_feedback
                        
                        try:
                            # Add comment to the element
                            annotation = add_comment_to_odt_element(doc, element['location']['element'], content_feedback)
                            if annotation is None:
                                # If annotation couldn't be added, increment the error count
                                comment_errors += 1
                                if verbose or comment_errors <= 3:  # Only show the first few errors
                                    print(f"\nCouldn't add comment to {element_type} {display} - falling back to summary document only")
                                elif comment_errors == 4:
                                    print("\nMultiple comment errors occurred. Suppressing further error messages...")
                        except Exception as comment_error:
                            comment_errors += 1
                            if verbose or comment_errors <= 3:  # Only show the first few errors
                                print(f"\nError adding comment to {element_type} {display}: {str(comment_error)}")
                            elif comment_errors == 4:
                                print("\nMultiple comment errors occurred. Suppressing further error messages...")
                        
                except Exception as para_e:
                    if verbose:
                        print(f"\nError analyzing {element_type} {display}: {str(para_e)}")
        
        # Clear progress bar line
        if not verbose:
            sys.stdout.write("\r" + " " * 80 + "\r")
            sys.stdout.flush()
        
        # Save the commented document
        doc.save(commented_path)
        
        # Create and save summary document if requested
        summary_path = None
        if summary_doc:
            summary_doc_obj = create_odt_summary_doc(text_elements, overall_feedback, model_name)
            summary_path = os.path.join(os.path.dirname(odt_path),
                                      f"{name_without_ext}_summary_{model_short_name}.odt")
            summary_doc_obj.save(summary_path)
        
        print(f"\n✅ Document analyzed with {feedback_count} comments added")
        if comment_errors > 0:
            print(f"⚠️  {comment_errors} comments could not be added due to technical limitations")
            print("    The summary document contains all feedback items.")
        
        print(f"✅ Commented document created: {commented_path}")
        if summary_path:
            print(f"✅ Summary document created: {summary_path}")
        print(f"\nNote: For best viewing of comments, open the document in LibreOffice Writer.")
        
        return commented_path, summary_path
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main() 