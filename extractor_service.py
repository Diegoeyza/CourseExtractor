import pdfplumber
import re
import os

class CourseExtractor:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.full_text = ""
        self.pages_text = []
        self.raw_tables = []
        self.pages_words = []
        
    def load_pdf(self):
        with pdfplumber.open(self.pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                self.pages_text.append(text)
                self.full_text += text + "\n"
                
                # Store words for layout analysis
                self.pages_words.append(page.extract_words())
                
                # Store tables with page index
                tables = page.extract_tables()
                self.raw_tables.append((page.page_number, tables))

    def extract_header(self):
        """Extracts Course Name, Area, Code, and Full NRC."""
        lines = [l.strip() for l in self.full_text.split('\n') if l.strip()]
        
        area = None
        code = None
        course_name = None
        
        m = re.search(r"Código:\s*([A-Z]+)\s*(\d+)", self.full_text)
        if m:
            area = m.group(1)
            code = m.group(2)
            
        first_line = lines[0] if lines else ""
        if " - " in first_line:
            course_name_raw = first_line.split(" - ")[0].strip()
        else:
            m_title = re.match(r"^(.*?)\s*-\s*\d+", self.full_text, re.DOTALL)
            if m_title:
                course_name_raw = m_title.group(1).strip()
            else:
                course_name_raw = first_line
        
        if area and code:
            course_name = course_name_raw.replace(f"{area} {code}", "").strip()
        else:
            course_name = course_name_raw
            
        return {
            "course_name": course_name,
            "area": area,
            "code": code,
            "full_nrc": f"{area}-{code}" if area and code else None
        }

    def extract_requirements(self):
        """Extracts requirements in format Name (NRC)."""
        pattern = r"Requisitos\s*/\s*Aprendizajes\s*previos:\s*(.*?)(?=Información de la asignatura|Tipo de asignatura|\n\n|\r\n\r\n)"
        m = re.search(pattern, self.full_text, re.IGNORECASE | re.DOTALL)
        if not m:
            return []
            
        req_str = m.group(1).replace('\n', ' ').strip()
        reqs = [r.strip() for r in req_str.split(',') if r.strip()]
        
        results = []
        for r in reqs:
            m_req = re.search(r"^(.*?)\s*\((.*?)\)$", r)
            if m_req:
                results.append({"name": m_req.group(1).strip(), "nrc": m_req.group(2).strip()})
            else:
                results.append({"name": r, "nrc": None})
        return results

    def extract_description(self):
        """Extracts description between header and APE."""
        pattern = r"Descripción de la asignatura\s*(.*?)(?=\s*Aporte al Perfil de Egreso)"
        m = re.search(pattern, self.full_text, re.IGNORECASE | re.DOTALL)
        if m:
            desc = m.group(1).strip()
            return desc
        return ""

    def extract_apes_fallback(self):
        """Fallback for when APE is table-like but lacks physical lines.
        Uses world coordinates to group text blocks correctly."""
        ape_items = []
        
        # 1. Identify which page has the APE section
        for pnum, text in enumerate(self.pages_text):
            if "ID_APE" in text and "Descripción de APE" in text:
                words = self.pages_words[pnum]
                
                # Find the vertical start of the section (bottom of the header line)
                header_bottom = 0
                for w in words:
                    if "ID_APE" in w['text'] or "Descripción" in w['text']:
                        header_bottom = max(header_bottom, w['bottom'])
                
                if header_bottom == 0: continue
                
                # Find all APE labels (APE 1, APE 2, etc.)
                ape_labels = []
                for i, w in enumerate(words):
                    if w['top'] >= header_bottom and re.match(r"^APE$", w['text'], re.IGNORECASE):
                        if i+1 < len(words) and re.match(r"^\d+$", words[i+1]['text']):
                            ape_labels.append({
                                "id": f"APE {words[i+1]['text']}", 
                                "top": w['top'], 
                                "bottom": w['bottom']
                            })
                
                if not ape_labels: continue
                
                # Group all description words (x0 > 200)
                desc_words = [w for w in words if w['top'] >= header_bottom and w['x0'] > 200]
                
                # For each APE, collect words that belong to it
                for i, label in enumerate(ape_labels):
                    y_min = 0
                    if i == 0:
                        y_min = header_bottom
                    else:
                        y_min = (ape_labels[i-1]['bottom'] + label['top']) / 2
                        
                    y_max = 0
                    if i + 1 < len(ape_labels):
                        y_max = (label['bottom'] + ape_labels[i+1]['top']) / 2
                    else:
                        y_max = 2000
                        
                    current_ape_words = [w for w in desc_words if w['top'] >= y_min and w['top'] < y_max]
                    
                    full_desc = " ".join([w['text'] for w in current_ape_words])
                    # Clean up: remove header fragments if any leaked
                    full_desc = re.sub(r"^(de APE \(aporte al perfil de egreso\)\s*)", "", full_desc, flags=re.IGNORECASE)
                    # Clean up artifacts like page numbers
                    full_desc = re.sub(r"Page \d+ of \d+", "", full_desc)
                    full_desc = re.sub(r"\s+", " ", full_desc).strip()
                    
                    if full_desc:
                        ape_items.append({"id": label['id'], "description": full_desc})
                
                # If we found APEs on a page, we stop looking at other pages
                if ape_items: break
                
        return ape_items

    def extract_tables(self):
        """Extracts APE and RA tables."""
        ape_items = []
        ra_items = []
        
        stop_general_ra = False
        current_mode = None

        for pnum, tables in self.raw_tables:
            if not tables: continue
            
            for table in tables:
                if not table or len(table) < 1: continue
                if stop_general_ra:
                    break
                
                for row in table:
                    row_text = " ".join([str(cell).lower() for cell in row if cell])
                    
                    if "descripción de contenidos por unidad" in row_text or "nombre de la unidad" in row_text:
                        stop_general_ra = True
                        break
                        
                    if "id_ape" in row_text or ("descripción" in row_text and "perfil de egreso" in row_text):
                        current_mode = "ape"
                        continue
                        
                    if ("id_ra" in row_text or ("resultados de aprendizaje" in row_text and "perfil de egreso" not in row_text)) \
                       and not ("contenidos" in row_text or "unidad" in row_text or "módulo" in row_text):
                        current_mode = "ra"
                        continue
                        
                    # parse based on mode
                    if current_mode:
                        cells = [str(c).strip() for c in row if c and str(c).strip()]
                        if len(cells) >= 2:
                            id_val = cells[0]
                            desc = cells[1]
                            
                            if current_mode == "ape":
                                if "id_ape" in id_val.lower() or "descripción" in id_val.lower(): continue
                                # Ensure it doesn't leak to general text (an ID shouldn't be a huge paragraph)
                                if len(id_val) < 20: 
                                    ape_items.append({"id": id_val, "description": desc.replace('\n', ' ')})
                                    
                            elif current_mode == "ra":
                                if "id_ra" in id_val.lower() or "resultado" in id_val.lower(): continue
                                if len(id_val) < 20:
                                    ra_items.append({"id": id_val, "description": desc.replace('\n', ' ')})
                                    
            if stop_general_ra:
                break

        # Remove duplicates if any
        ape_items = [dict(t) for t in {tuple(d.items()) for d in ape_items}]
        ra_items = [dict(t) for t in {tuple(d.items()) for d in ra_items}]

        # If no APEs found via tables, try regex fallback
        if not ape_items:
            print("DEBUG: No APE tables found. Attempting regex fallback...")
            ape_items = self.extract_apes_fallback()
            if ape_items:
                print(f"DEBUG: Fallback found {len(ape_items)} APEs.")
                                
        return ape_items, ra_items

    def extract_bibliography(self):
        """Extracts bibliography between markers."""
        start_marker = r"Recursos de Aprendizaje - Bibliografía Básica"
        # The marker might have an accent or not
        end_marker = r"Recursos de Aprendizaje - Bibliograf[íi]a Complementaria"
        
        pattern = f"{start_marker}(.*?)(?:{end_marker}|Ausencia a Evaluaciones|Integridad Académica|$)"
        m = re.search(pattern, self.full_text, re.IGNORECASE | re.DOTALL)
            
        if not m:
            return []
            
        text = m.group(1).strip()
        
        # Clean up Page X of Y
        text = re.sub(r"Page \d+ of \d+", "", text)
        
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        
        results = []
        current_book = None
        
        metadata_fields = ["ISBN:", "Autor:", "Editor:", "Fecha de publicación:"]
        
        for i, line in enumerate(lines):
            is_metadata = any(line.startswith(field) for field in metadata_fields)
            
            if is_metadata:
                if current_book:
                    current_book["metadata"].append(line)
                elif i > 0:
                    # Check if the title was the previous line (and it wasn't a metadata field itself)
                    prev_line = lines[i-1]
                    is_prev_metadata = any(prev_line.startswith(field) for field in metadata_fields)
                    if not is_prev_metadata:
                        current_book = {"title": prev_line, "metadata": [line]}
                        results.append(current_book)
            else:
                # If we current have a book, this might be a multi-line metadata or a new title.
                # In the current format, titles are single lines before ISBN.
                # If current_book exists, we stop collecting for it if this isn't metadata.
                current_book = None
                
        # Final formatting: join metadata into single string
        final_results = []
        for res in results:
            final_results.append({
                "title": res["title"],
                "metadata": " | ".join(res["metadata"])
            })
            
        return final_results

    def get_structured_data(self):

        """Returns structured data suitable for database models."""
        self.load_pdf()
        header = self.extract_header()
        requirements = self.extract_requirements()
        description = self.extract_description()
        apes, ras = self.extract_tables()
        bibliography = self.extract_bibliography()
        
        return {
            "course": {
                "name": header["course_name"],
                "area": header["area"],
                "code": header["code"],
                "full_nrc": header["full_nrc"],
                "description": description
            },
            "prerequisites": requirements,
            "apes": apes,
            "ras": ras,
            "bibliography": bibliography
        }

