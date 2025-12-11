#!/usr/bin/env python3
"""
Patient Invoice & Cover Letter Generator
Generates per-patient invoices (PDF) and cover letters (DOCX) from patient roster and invoice data.
"""

import pandas as pd
import numpy as np
from pathlib import Path
import json
import re
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional
import logging
from dataclasses import dataclass
import os
from difflib import SequenceMatcher

# PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER

# DOCX handling
from docx import Document

@dataclass
class PatientData:
    """Patient information from roster"""
    prn: str
    first_name: str
    last_name: str
    dob: str
    address_line1: str
    address_line2: str
    city: str
    state: str
    postal_code: str
    
@dataclass
class InvoiceLine:
    """Individual service line from invoice"""
    service_date: str
    description: str
    amount: float
    is_previous_balance: bool = False

@dataclass
class ProcessingSummary:
    """Summary of processing results"""
    processed_patients: List[str]
    skipped_patients: List[Tuple[str, str]]  # (name, reason)
    errors: List[Tuple[str, str]]  # (patient_name, error)
    total_processed: int
    total_skipped: int
    total_errors: int
    total_amount_due: float
    processing_date: str

class PatientInvoiceGenerator:
    """Main class for generating patient invoices and cover letters"""
    
    def __init__(self, amount_due_strategy: str = "auto", statement_date: Optional[str] = None):
        self.amount_due_strategy = amount_due_strategy
        self.statement_date = datetime.strptime(statement_date, "%Y-%m-%d") if statement_date else datetime.now()
        self.payment_due_date = self._calculate_payment_due_date()
        
        # Column mapping for normalization - UPDATED
        self.column_aliases = {
            'name': ["Name", "Patient Name", "[LastName, FirstName]"],
            'visit_date': ["Visit Date", "Service Date", "DOS", "Date of Service"],
            'total_amount': ["Total amount", "Charge", "Billed Amount", "Total Charge"],
            'copay': ["Copay", "Co-pay", "Copayment"],
            'paid': ["Paid", "Patient Paid", "Payments"],
            'previous_balance': ["Previous Balance", "Outstanding Balance", "Prior Balance", "Carryover"],
            'insurance': ["Insurance"],
            'type_of_service': ["Type Of Service", "Service Type", "Description", "Service Description"]  # NEW
        }
        
        # Setup logging
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        self.logger = logging.getLogger(__name__)
        
    def _calculate_payment_due_date(self) -> datetime:
        """Calculate payment due date (1 month from statement date, adjusted for weekends)"""
        due_date = self.statement_date + timedelta(days=30)
        
        # If weekend, move to next Monday
        if due_date.weekday() == 5:  # Saturday
            due_date += timedelta(days=2)
        elif due_date.weekday() == 6:  # Sunday
            due_date += timedelta(days=1)
            
        return due_date
    
    def _normalize_column_name(self, df: pd.DataFrame, target_column: str, custom_mapping: Dict = None) -> Optional[str]:
        """Find the actual column name using aliases"""
        if custom_mapping and target_column in custom_mapping:
            if custom_mapping[target_column] in df.columns:
                return custom_mapping[target_column]
        
        if target_column in self.column_aliases:
            for alias in self.column_aliases[target_column]:
                if alias in df.columns:
                    return alias
        
        # Direct match
        if target_column in df.columns:
            return target_column
            
        return None
    
    def _calculate_string_similarity(self, str1: str, str2: str) -> float:
        """Calculate similarity between two strings using sequence matching"""
        if not str1 or not str2:
            return 0.0
        
        # Normalize strings
        s1 = str1.lower().strip()
        s2 = str2.lower().strip()
        
        if s1 == s2:
            return 1.0
        
        # Use SequenceMatcher for similarity
        similarity = SequenceMatcher(None, s1, s2).ratio()
        
        # Boost score for very close matches (like "Shley" vs "Shely")
        if similarity >= 0.8:
            return similarity
        
        # Check for substring matches
        if s1 in s2 or s2 in s1:
            return max(similarity, 0.7)
        
        # Check for common prefixes/suffixes
        if len(s1) >= 3 and len(s2) >= 3:
            if s1[:3] == s2[:3] or s1[-3:] == s2[-3:]:
                return max(similarity, 0.6)
        
        return similarity
    
    def _sanitize_filename(self, name: str) -> str:
        """Sanitize filename to only contain A-Z, a-z, 0-9, _"""
        return re.sub(r'[^A-Za-z0-9_]', '_', name)
    
    def _parse_patient_name(self, name: str) -> Tuple[str, str]:
        """Parse 'LastName, FirstName' format with support for complex names"""
        if ',' in name:
            parts = name.split(',', 1)
            last_name = parts[0].strip()
            first_name = parts[1].strip()
            
            # Handle complex last names like "Russell (Kwon)"
            # Remove parentheses and extra names for cleaner matching
            if '(' in last_name:
                # Extract the main last name before parentheses
                main_last_name = last_name.split('(')[0].strip()
                return first_name, main_last_name
            
            return first_name, last_name
        else:
            # Fallback: assume single name is last name
            return "", name.strip()
    
    def load_patient_roster(self, roster_file: str) -> Dict[str, PatientData]:
        """Load patient roster CSV and create lookup dictionary"""
        try:
            df = pd.read_csv(roster_file)
            self.logger.info(f"Loaded roster with {len(df)} patients")
            
            # Print first few rows to understand the structure
            self.logger.info(f"Roster columns: {list(df.columns)}")
            if len(df) > 0:
                self.logger.info(f"Sample row: {df.iloc[0].to_dict()}")
            
            patients = {}
            
            # Try to identify columns dynamically
            possible_prn_cols = ['Patient Record Number', 'PRN', 'ID', 'Patient ID']
            possible_first_cols = ['First name', 'First Name', 'FirstName', 'Given Name']
            possible_last_cols = ['Last name', 'Last Name', 'LastName', 'Surname', 'Family Name']
            possible_dob_cols = ['DOB', 'Date of Birth', 'Birth Date']
            possible_addr1_cols = ['Address Line 1', 'Address 1', 'Street Address', 'Address']
            possible_addr2_cols = ['Address Line 2', 'Address 2', 'Apt', 'Suite']
            possible_city_cols = ['City']
            possible_state_cols = ['State', 'ST']
            possible_zip_cols = ['Postal Code', 'Zip Code', 'ZIP', 'Zip']
            
            def find_column(df, possible_names):
                for name in possible_names:
                    if name in df.columns:
                        return name
                return None
            
            # Standard CSV processing
            prn_col = find_column(df, possible_prn_cols)
            first_col = find_column(df, possible_first_cols) 
            last_col = find_column(df, possible_last_cols)
            dob_col = find_column(df, possible_dob_cols)
            addr1_col = find_column(df, possible_addr1_cols)
            addr2_col = find_column(df, possible_addr2_cols)
            city_col = find_column(df, possible_city_cols)
            state_col = find_column(df, possible_state_cols)
            zip_col = find_column(df, possible_zip_cols)
            
            # Log what columns were found
            self.logger.info(f"Detected columns - PRN: {prn_col}, First: {first_col}, Last: {last_col}")
            
            for _, row in df.iterrows():
                # Clean up postal code to remove decimal formatting
                postal_code = row.get(zip_col, '') if zip_col else ''
                if pd.notnull(postal_code):
                    # Convert to string and remove .0 if present
                    postal_code = str(postal_code)
                    if postal_code.endswith('.0'):
                        postal_code = postal_code[:-2]
                    # Also remove any other decimal patterns
                    if '.' in postal_code and postal_code.replace('.', '').isdigit():
                        postal_code = postal_code.split('.')[0]
                else:
                    postal_code = ''
                
                patient = PatientData(
                    prn=str(row.get(prn_col, '')) if prn_col and pd.notnull(row.get(prn_col)) else '',
                    first_name=str(row.get(first_col, '')) if first_col and pd.notnull(row.get(first_col)) else '',
                    last_name=str(row.get(last_col, '')) if last_col and pd.notnull(row.get(last_col)) else '',
                    dob=str(row.get(dob_col, '')) if dob_col and pd.notnull(row.get(dob_col)) else '',
                    address_line1=str(row.get(addr1_col, '')) if addr1_col and pd.notnull(row.get(addr1_col)) else '',
                    address_line2=str(row.get(addr2_col, '')) if addr2_col and pd.notnull(row.get(addr2_col)) else '',
                    city=str(row.get(city_col, '')) if city_col and pd.notnull(row.get(city_col)) else '',
                    state=str(row.get(state_col, '')) if state_col and pd.notnull(row.get(state_col)) else '',
                    postal_code=postal_code
                )
                
                # Create lookup keys
                if patient.prn and patient.prn != 'nan':
                    patients[f"prn_{patient.prn}"] = patient
                
                name_key = f"{patient.first_name.lower()}_{patient.last_name.lower()}"
                if name_key not in patients:  # First match wins
                    patients[name_key] = patient
            
            self.logger.info(f"Loaded {len(patients)} unique patients into lookup")
            return patients
            
        except Exception as e:
            self.logger.error(f"Error loading patient roster: {e}")
            raise
    
    # Add this helper method to format dates consistently
    def _format_date_for_display(self, date_value) -> str:
        """Format date value to MM/DD/YYYY format"""
        if pd.isna(date_value) or date_value == '' or date_value is None:
            return ''
        
        try:
            # Convert to string first
            date_str = str(date_value)
            
            # Try to parse as datetime if it's not already
            if isinstance(date_value, str):
                # Handle different possible input formats
                for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y', '%Y/%m/%d']:
                    try:
                        parsed_date = datetime.strptime(date_str.split()[0], fmt)  # Take only date part
                        return parsed_date.strftime('%m/%d/%Y')
                    except ValueError:
                        continue
            elif hasattr(date_value, 'strftime'):
                # If it's already a datetime object
                return date_value.strftime('%m/%d/%Y')
            
            # If we can't parse it, try to extract date from string
            date_part = date_str.split()[0] if ' ' in date_str else date_str
            
            # Try pandas to_datetime as last resort
            parsed_date = pd.to_datetime(date_part, errors='coerce')
            if not pd.isna(parsed_date):
                return parsed_date.strftime('%m/%d/%Y')
            
            # If all else fails, return the original string
            return date_part
            
        except Exception as e:
            # Log the error but don't break the process
            self.logger.warning(f"Could not format date '{date_value}': {e}")
            return str(date_value) if date_value else ''

    # Update the load_invoice_data method to include the new column
    def load_invoice_data(self, invoice_file: str, custom_mapping: Dict = None) -> pd.DataFrame:
        """Load invoice Excel file and normalize columns"""
        try:
            df = pd.read_excel(invoice_file, sheet_name='Sheet1')
            self.logger.info(f"Loaded invoice data from Sheet1 with {len(df)} rows")
            
            # Normalize column names
            column_mapping = {}
            required_columns = ['name', 'visit_date', 'total_amount', 'copay', 'paid']
            
            for col in required_columns:
                actual_col = self._normalize_column_name(df, col, custom_mapping)
                if actual_col is None:
                    raise ValueError(f"Required column '{col}' not found. Available columns: {list(df.columns)}")
                column_mapping[actual_col] = col
            
            # Optional columns - UPDATED
            optional_columns = ['previous_balance', 'insurance', 'type_of_service']
            for col in optional_columns:
                actual_col = self._normalize_column_name(df, col, custom_mapping)
                if actual_col:
                    column_mapping[actual_col] = col
            
            # Rename columns
            df = df.rename(columns=column_mapping)
            
            # Fill missing optional columns
            if 'previous_balance' not in df.columns:
                df['previous_balance'] = 0
            if 'insurance' not in df.columns:
                df['insurance'] = ''
            if 'type_of_service' not in df.columns:  # NEW
                df['type_of_service'] = ''
            
            # Clean and convert data types
            df['total_amount'] = pd.to_numeric(df['total_amount'], errors='coerce').fillna(0)
            df['copay'] = pd.to_numeric(df['copay'], errors='coerce').fillna(0)
            df['paid'] = pd.to_numeric(df['paid'], errors='coerce').fillna(0)
            df['previous_balance'] = pd.to_numeric(df['previous_balance'], errors='coerce').fillna(0)
            
            # Clean type_of_service column - NEW
            df['type_of_service'] = df['type_of_service'].fillna('').astype(str)
            
            return df
            
        except Exception as e:
            self.logger.error(f"Error loading invoice data: {e}")
            raise
    
    def _calculate_amount_due(self, row: pd.Series) -> float:
        """Calculate amount due based on strategy"""
        total_amount = float(row['total_amount'])
        copay = float(row['copay'])
        paid = float(row['paid'])
        
        if self.amount_due_strategy == "auto":
            if copay > 0:
                amount_due = copay - paid
            else:
                amount_due = total_amount - paid
        elif self.amount_due_strategy == "copay_minus_paid":
            amount_due = copay - paid
        elif self.amount_due_strategy == "total_minus_paid":
            amount_due = total_amount - paid
        else:
            raise ValueError(f"Unknown amount due strategy: {self.amount_due_strategy}")
        
        return max(0, amount_due)  # Floor at 0
    
    def _match_patient(self, name: str, patients: Dict[str, PatientData]) -> Tuple[Optional[PatientData], bool]:
        """Match patient by name with fuzzy matching support"""
        first_name, last_name = self._parse_patient_name(name)
        
        # Try exact match first
        name_key = f"{first_name.lower()}_{last_name.lower()}"
        if name_key in patients:
            return patients[name_key], False
        
        # Try partial/fuzzy matching with string similarity
        first_lower = first_name.lower().strip()
        last_lower = last_name.lower().strip()
        
        # Extract name components for better matching
        first_parts = [part.strip() for part in first_lower.replace(',', ' ').split() if part.strip()]
        last_parts = [part.strip() for part in last_lower.replace(',', ' ').split() if part.strip()]
        
        matches = []
        
        for key, patient in patients.items():
            if key.startswith('prn_'):
                continue
                
            patient_first = patient.first_name.lower().strip()
            patient_last = patient.last_name.lower().strip()
            
            # Calculate similarity scores
            first_name_score = 0
            last_name_score = 0
            
            # Score first name matching
            if first_parts and patient_first:
                max_first_score = 0
                for first_part in first_parts:
                    # Direct similarity
                    sim = self._calculate_string_similarity(first_part, patient_first)
                    max_first_score = max(max_first_score, sim)
                    
                    # Also check against individual words in patient name
                    for patient_word in patient_first.split():
                        sim = self._calculate_string_similarity(first_part, patient_word)
                        max_first_score = max(max_first_score, sim)
                
                first_name_score = max_first_score
            
            # Score last name matching
            if last_parts and patient_last:
                max_last_score = 0
                for last_part in last_parts:
                    # Direct similarity
                    sim = self._calculate_string_similarity(last_part, patient_last)
                    max_last_score = max(max_last_score, sim)
                    
                    # Also check against individual words in patient name
                    for patient_word in patient_last.split():
                        sim = self._calculate_string_similarity(last_part, patient_word)
                        max_last_score = max(max_last_score, sim)
                
                last_name_score = max_last_score
            
            # Calculate overall match score
            # Both names need decent scores to be considered a match
            if first_name_score >= 0.6 and last_name_score >= 0.6:
                overall_score = (first_name_score + last_name_score) / 2
                matches.append((patient, overall_score, first_name_score, last_name_score))
        
        if matches:
            # Sort by best overall match score
            matches.sort(key=lambda x: x[1], reverse=True)
            best_match = matches[0]
            
            # Log the fuzzy match with detailed scores
            self.logger.info(f"Fuzzy match found for '{name}': "
                           f"{best_match[0].first_name} {best_match[0].last_name} "
                           f"(PRN: {best_match[0].prn}) - Overall: {best_match[1]:.1%}, "
                           f"First: {best_match[2]:.1%}, Last: {best_match[3]:.1%}")
            
            # Return ambiguous if multiple high-scoring matches
            is_ambiguous = len([m for m in matches if m[1] >= 0.85]) > 1
            
            return best_match[0], is_ambiguous
        
        # No match found
        self.logger.warning(f"No patient match found for: {name}")
        return None, False
    
    # Update the _generate_invoice_lines method to use the service type
    def _generate_invoice_lines(self, patient_df: pd.DataFrame) -> Tuple[List[InvoiceLine], float, pd.DataFrame]:
        """Generate invoice lines for a patient"""
        lines = []
        
        # Handle previous balance (take from first row)
        first_row = patient_df.iloc[0]
        previous_balance = float(first_row.get('previous_balance', 0))
        
        if previous_balance > 0:
            lines.append(InvoiceLine(
                service_date="",
                description="Previous Balance", 
                amount=previous_balance,
                is_previous_balance=True
            ))
        
        # Process service lines - UPDATED
        for _, row in patient_df.iterrows():
            amount_due = self._calculate_amount_due(row)
            
            if amount_due > 0:  # Only include lines with amounts due
                # Get service type from the new column, use default if empty
                service_type = row.get('type_of_service', '').strip()
                if not service_type:
                    service_type = "Psychotherapy and/or Med Management"  # Default
                
                lines.append(InvoiceLine(
                    service_date=str(row['visit_date']),
                    description=service_type,  # Use the actual service type
                    amount=amount_due,
                    is_previous_balance=False
                ))
        
        # Calculate total_due as subtotal_copay minus subtotal_paid
        subtotal_copay = float(patient_df['copay'].sum()) + previous_balance
        subtotal_paid = float(patient_df['paid'].sum())
        total_due = max(0, subtotal_copay - subtotal_paid)  # Floor at 0
        
        return lines, total_due, patient_df
    
    # Update the _generate_pdf_invoice method to use dynamic descriptions
    def _generate_pdf_invoice(self, patient: PatientData, lines: List[InvoiceLine], 
                            total_due: float, patient_df: pd.DataFrame, output_path: Path):
        """Generate PDF invoice"""
        try:
            # Count items to determine layout optimization level
            num_items = len(patient_df[
                (patient_df.get('paid', 0) != 0) | 
                (patient_df.get('copay', 0) != 0)
            ])
            
            # Check for previous balance
            has_previous_balance = any(line.is_previous_balance for line in lines)
            total_rows = num_items + (1 if has_previous_balance else 0) + 3  # +3 for header, subtotal, total
            
            # Determine optimization level
            if total_rows <= 18:
                # Standard layout - no changes needed
                font_header = 12
                font_header2 = 10
                font_title = 13
                font_body = 10
                font_table = 9
                row_height = None  # Default
                top_margin = 0.5
                bottom_margin = 0.75
                spacer_header = 8
                spacer_sections = 20
                spacer_small = 20
                footer_font = 12
            elif total_rows <= 25:
                # Slight compression
                font_header = 11
                font_header2 = 9
                font_title = 12
                font_body = 9
                font_table = 8
                row_height = 14
                top_margin = 0.4
                bottom_margin = 0.6
                spacer_header = 6
                spacer_sections = 15
                spacer_small = 15
                footer_font = 10
            elif total_rows <= 35:
                # Moderate compression
                font_header = 10
                font_header2 = 8
                font_title = 11
                font_body = 8
                font_table = 7
                row_height = 12
                top_margin = 0.35
                bottom_margin = 0.5
                spacer_header = 4
                spacer_sections = 10
                spacer_small = 10
                footer_font = 9
            else:
                # Maximum compression
                font_header = 9
                font_header2 = 7
                font_title = 10
                font_body = 7
                font_table = 6
                row_height = 10
                top_margin = 0.3
                bottom_margin = 0.4
                spacer_header = 3
                spacer_sections = 8
                spacer_small = 8
                footer_font = 8

            # Make footer font size available to footer callback
            self.footer_font = footer_font
        
            # Create document with optimized margins
            doc = SimpleDocTemplate(str(output_path), pagesize=letter,
                                topMargin=top_margin*inch, bottomMargin=bottom_margin*inch,
                                leftMargin=0.75*inch, rightMargin=0.75*inch)
            
            story = []
            styles = getSampleStyleSheet()
            
            # Custom styles
            header_style = ParagraphStyle(
                'Header',
                parent=styles['Normal'],
                fontSize=font_header,  # Changed to dynamic
                alignment=TA_CENTER,
                spaceAfter=3,
                fontName='Helvetica-Bold'
            )

            header_2_style = ParagraphStyle(
                'Header2',  # Fixed duplicate name
                parent=styles['Normal'],
                fontSize=font_header2,  # Changed to dynamic
                alignment=TA_CENTER,
                spaceAfter=2,
                fontName='Helvetica-Bold'
            )
            
            title_style = ParagraphStyle(
                'Title',
                parent=styles['Normal'],
                fontSize=font_title,  # Changed to dynamic
                alignment=TA_CENTER,
                spaceAfter=spacer_sections,  # Changed to dynamic
                fontName='Helvetica-Bold'
            )
            
            # Header section
            story.append(Paragraph("ACCESS MULTI-SPECIALTY MEDICAL CLINIC, INC.", header_style))
            story.append(Paragraph("MICHAEL U. LEVINSON, MD, PH D.", header_style))
            story.append(Paragraph("BOARD CERTIFIED PSYCHIATRIST", header_style))
            story.append(Spacer(1, spacer_header))  # Instead of 8
            
            story.append(Paragraph("OFFICE ADDRESS: 25 EDWARDS COURT, SUITE 101, BURLINGAME, CA 94010", header_2_style))
            story.append(Paragraph("MAILING ADDRESS: PO BOX 351, BURLINGAME, CA 94011", header_2_style))
            story.append(Paragraph("EMAIL: ACCESS.MSMC@GMAIL.COM", header_2_style))
            story.append(Spacer(1, spacer_sections))  # Instead of 20
            
            # Patient information and payment instructions side by side
            display_postal = patient.postal_code
            if display_postal and '.' in display_postal:
                display_postal = display_postal.split('.')[0]

            # Build patient info content
            patient_info_text = f"{patient.first_name.upper()} {patient.last_name.upper()}\n"
            if patient.address_line1:
                patient_info_text += f"{patient.address_line1}\n"
            if patient.address_line2:
                patient_info_text += f"{patient.address_line2}\n"
            patient_info_text += f"{patient.city}, {patient.state} {display_postal}"

            # Build payment instructions content
            payment_info_text = ("Please note we do not accept credit cards.\n"
                                "1. Zelle access.msmc@gmail.com (IRA Billing and Mgmt)\n"
                                "2. Check payable to: Michael Levinson, MD\n"
                                "   PO Box 351, Burlingame, CA 94011")

            # Create side-by-side table
            combined_info = [
                [patient_info_text, payment_info_text]
            ]

            combined_table = Table(combined_info, colWidths=[3*inch, 3*inch])
            combined_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
                ('FONTNAME', (1, 0), (1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), font_body),  # Dynamic
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ]))

            story.append(combined_table)
            story.append(Spacer(1, 20))
            
            # Title
            story.append(Paragraph("PATIENT STATEMENT", title_style))
            
            # Statement info section
            statement_info = [
                ["STATEMENT DATE:", "Payment due date:"],
                [self.statement_date.strftime('%m/%d/%Y'), self.payment_due_date.strftime('%m/%d/%Y')]
            ]
            
            statement_table = Table(statement_info, colWidths=[2*inch, 2*inch])
            statement_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), font_body),  # Dynamic
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]))
            
            story.append(statement_table)
            story.append(Spacer(1, 20))
            
# Replace the entire table generation section in _generate_pdf_invoice with this corrected version:

            # Service details table - CORRECTED VERSION
            table_data = [['Service Date(s)', 'Description', 'Amount Paid', 'Copay/Deductible']]

            total_paid = 0
            total_copay = 0

            # First, handle any previous balance
            previous_balance_added = False
            for line in lines:
                if line.is_previous_balance and not previous_balance_added:
                    table_data.append([
                        '',
                        'Previous Balance',
                        '$ -',
                        f'$ {line.amount:.2f}'
                    ])
                    total_copay += line.amount
                    previous_balance_added = True
                    break

            # Then process all invoice rows to show services and payments
            for _, row in patient_df.iterrows():
                visit_date_raw = row['visit_date']
                paid_amount = float(row.get('paid', 0))
                copay_amount = float(row.get('copay', 0))
                
                # Show all rows that have any activity
                amount_due = self._calculate_amount_due(row)
                
                if amount_due > 0 or paid_amount > 0 or copay_amount > 0:
                    # Format the date for display as MM/DD/YYYY
                    display_date = self._format_date_for_display(visit_date_raw)
                    
                    # Get service description
                    service_type = row.get('type_of_service', '').strip()
                    if not service_type:
                        service_type = "Mental Health Visit"  # Default for table display
                    
                    # Add single row to table - FIXED
                    table_data.append([
                        display_date,
                        service_type,
                        f'$ {paid_amount:.2f}' if paid_amount > 0 else '$ -',
                        f'$ {copay_amount:.2f}' if copay_amount > 0 else '$ -'
                    ])
                    
                    total_paid += paid_amount
                    total_copay += copay_amount

            # Add subtotal row
            table_data.append(['', 'SUBTOTAL', f'$ {total_paid:.2f}', f'$ {total_copay:.2f}'])
            table_data.append(['', 'TOTAL', '', f'$ {total_due:.2f}'])
            
            # Debug: Print table structure to see what's being generated
            self.logger.info(f"Table data structure: {len(table_data)} rows")
            for i, row in enumerate(table_data):
                self.logger.info(f"Row {i}: {row}")
            
            # Create table with dynamic row heights
            if row_height:
                # Set specific row heights for all rows
                row_heights = [row_height + 2] + [row_height] * (len(table_data) - 1)
                service_table = Table(table_data, 
                                    colWidths=[1.5*inch, 2.5*inch, 1.2*inch, 1.3*inch],
                                    rowHeights=row_heights)
            else:
                # Use default row heights
                service_table = Table(table_data, 
                                    colWidths=[1.5*inch, 2.5*inch, 1.2*inch, 1.3*inch])
            
            service_table.setStyle(TableStyle([
                # Header row
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), font_body),  # Dynamic header
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (2, 0), (3, -1), 'RIGHT'),
                
                # Data rows
                ('FONTNAME', (0, 1), (-1, -3), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -3), font_table),  # Dynamic data rows
                
                # Subtotal and total rows
                ('FONTNAME', (0, -2), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, -2), (-1, -1), font_body),  # Dynamic totals
                ('LINEABOVE', (0, -2), (-1, -2), 1, colors.black),
                ('LINEABOVE', (0, -1), (-1, -1), 2, colors.black),
                
                # Grid
                ('GRID', (0, 0), (-1, -3), 0.5, colors.black),
                ('BOX', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

                # Reduce padding when compressed
                ('TOPPADDING', (0, 1), (-1, -1), 1 if total_rows > 25 else 3),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 1 if total_rows > 25 else 3),
            ]))
            
            story.append(service_table)
            story.append(Spacer(1, spacer_small))  # After service table
            
            # Amount due section
            amount_section = [
                ["YOUR PORTION DUE:", "AMOUNT ENCLOSED:"],
                [f"${total_due:.2f}", ""]
            ]
            
            amount_table = Table(amount_section, colWidths=[2*inch, 2*inch])
            amount_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), font_body + 1),  # Slightly larger than body
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('BOX', (0, 0), (0, 1), 1, colors.black),
                ('BOX', (1, 0), (1, 1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            
            story.append(amount_table)
            story.append(Spacer(1, spacer_sections if total_rows <= 25 else spacer_small))  # After amount table
            
            # Provider signature
            signature_style = ParagraphStyle(
                'Signature',
                parent=styles['Normal'],
                fontSize=font_body + 1,  # Dynamic
                alignment=TA_RIGHT,
                fontName='Helvetica-Bold'
            )

            story.append(Paragraph("_________________________________", signature_style))
            story.append(Spacer(1, 12 if total_rows <= 25 else 8))  # After signature line
            story.append(Paragraph("Provider Signature - Michael Levinson, MD", signature_style))
            
            # Build the document and attach the footer method
            doc.build(
                story,
                onFirstPage=self.add_optimized_footer,
                onLaterPages=self.add_optimized_footer,
            )
            self.logger.info(f"Generated PDF invoice: {output_path}")

        except Exception as e:
            self.logger.error(f"Error generating PDF invoice: {e}")
            raise

    def add_optimized_footer(self, canvas, doc):
        """Add footer to each page with dynamic font size"""
        canvas.saveState()
        footer_text = "If you have questions regarding your bill, please contact us at (415)857-1151."
    
        # Fallback if, for some reason, footer_font wasn't set
        font_size = getattr(self, "footer_font", 10)
    
        canvas.setFont("Helvetica", font_size)
    
        # Center the text
        text_width = canvas.stringWidth(footer_text, "Helvetica", font_size)
        x_position = (letter[0] - text_width) / 2
        canvas.drawString(x_position, 0.5 * inch, footer_text)
        canvas.restoreState()
        
    
    def _generate_cover_letter(self, patient: PatientData, template_file: str, output_path: Path):
        """Generate cover letter DOCX from template"""
        try:
            # Check if file already exists and is potentially locked
            if output_path.exists():
                try:
                    # Try to delete the existing file
                    output_path.unlink()
                except OSError as e:
                    self.logger.warning(f"Could not delete existing file {output_path}: {e}")
                    # Try alternative filename
                    counter = 1
                    while True:
                        new_path = output_path.parent / f"{output_path.stem}_{counter}.docx"
                        if not new_path.exists():
                            output_path = new_path
                            break
                        counter += 1
                        if counter > 10:  # Prevent infinite loop
                            raise OSError(f"Cannot create unique filename for {output_path}")
            
            # Load template
            doc = Document(template_file)
            
            # Clean postal code for display
            display_postal = patient.postal_code
            if display_postal and '.' in display_postal:
                display_postal = display_postal.split('.')[0]
            
            # Replacement mappings - handle empty values gracefully
            replacements = {
                '[First Name]': patient.first_name or '',
                '[Last Name]': patient.last_name or '',
                '[Full Name]': f"{patient.first_name} {patient.last_name}".strip(),
                '[Address Line 1]': patient.address_line1 or '',
                '[Address Line 2]': patient.address_line2 or '',
                '[City]': patient.city or '',
                '[State]': patient.state or '',
                '[Postal Code]': display_postal or '',
                '[Patient Record Number]': patient.prn or ''
            }
            
            # More robust replacement that handles runs and formatting
            def replace_text_in_paragraph(paragraph, old_text, new_text):
                """Replace text while preserving formatting"""
                if old_text in paragraph.text:
                    # Handle simple case first
                    for run in paragraph.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)
                    
                    # Handle case where placeholder spans multiple runs
                    full_text = paragraph.text
                    if old_text in full_text:
                        # Clear all runs and create new one with replaced text
                        new_text_full = full_text.replace(old_text, new_text)
                        paragraph.clear()
                        paragraph.add_run(new_text_full)
            
            # Replace in paragraphs
            for paragraph in doc.paragraphs:
                for placeholder, value in replacements.items():
                    replace_text_in_paragraph(paragraph, placeholder, value)
            
            # Replace in tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for placeholder, value in replacements.items():
                                replace_text_in_paragraph(paragraph, placeholder, value)
            
            # Handle address formatting - clean up empty address line 2
            for paragraph in doc.paragraphs:
                # If address line 2 is empty, remove the extra comma and space
                text = paragraph.text
                if '[Address Line 2]' in text and not patient.address_line2:
                    # Clean up ", [Address Line 2]" when address line 2 is empty
                    text = text.replace(', [Address Line 2]', '')
                    text = text.replace('[Address Line 2], ', '')
                    text = text.replace('[Address Line 2]', '')
                    paragraph.clear()
                    paragraph.add_run(text)
            
            doc.save(str(output_path))
            self.logger.info(f"Generated cover letter: {output_path}")
            
        except Exception as e:
            self.logger.error(f"Error generating cover letter: {e}")
            raise

    def _generate_envelope_pdf(self, patient: PatientData, template_file: str, output_path: Path):
        """Generate envelope PDF sized for Com-10 envelopes (4.13" x 9.5") - single page"""
        try:
            # Com-10 envelope dimensions
            envelope_width = 9.5 * inch
            envelope_height = 4.13 * inch
            envelope_size = (envelope_width, envelope_height)
            
            doc = SimpleDocTemplate(str(output_path), pagesize=envelope_size,
                                topMargin=0.2*inch, bottomMargin=0.2*inch,
                                leftMargin=0.2*inch, rightMargin=0.2*inch)
            
            story = []
            styles = getSampleStyleSheet()
            
            # Return address style (top-left, very compact)
            return_address_style = ParagraphStyle(
                'ReturnAddress',
                parent=styles['Normal'],
                fontSize=10,
                alignment=TA_LEFT,
                spaceAfter=0,
                spaceBefore=0,
                fontName='Helvetica',
                leftIndent=0,
                leading=12  # Tight line spacing
            )
            
            return_address_bold_style = ParagraphStyle(
                'ReturnAddressBold',
                parent=styles['Normal'],
                fontSize=11,
                alignment=TA_LEFT,
                spaceAfter=1,
                spaceBefore=0,
                fontName='Helvetica-Bold',
                leftIndent=0,
                leading=12
            )
            
            # Delivery address style (center area, compact)
            delivery_address_style = ParagraphStyle(
                'DeliveryAddress',
                parent=styles['Normal'],
                fontSize=10,
                alignment=TA_LEFT,
                spaceAfter=0,
                spaceBefore=0,
                fontName='Helvetica',
                leftIndent=3.5*inch,  # Position for Com-10 envelope
                leading=12  # Tight line spacing
            )
            
            delivery_name_style = ParagraphStyle(
                'DeliveryName',
                parent=styles['Normal'],
                fontSize=11,
                alignment=TA_LEFT,
                spaceAfter=2,
                spaceBefore=0,
                fontName='Helvetica-Bold',
                leftIndent=3.5*inch,
                leading=12
            )
            
            # Return Address (very compact, top-left)
            story.append(Paragraph("<b>Access Multi-Specialty</b>", return_address_bold_style))
            story.append(Paragraph("<b>Medical Clinic, Inc.</b>", return_address_bold_style))
            story.append(Paragraph("PO Box 351", return_address_style))
            story.append(Paragraph("Burlingame, CA 94011", return_address_style))
            
            # Minimal spacing between return and delivery address
            story.append(Spacer(1, 1*inch))
            
            # Delivery Address (center-right)
            display_postal = patient.postal_code
            if display_postal and '.' in display_postal:
                display_postal = display_postal.split('.')[0]
            
            # Patient name
            patient_full_name = f"{patient.first_name} {patient.last_name}"
            story.append(Paragraph(f"<b>{patient_full_name}</b>", delivery_name_style))
            
            # Patient address
            if patient.address_line1:
                story.append(Paragraph(patient.address_line1, delivery_address_style))
            if patient.address_line2:
                story.append(Paragraph(patient.address_line2, delivery_address_style))
            
            # City, State ZIP
            city_state_zip = f"{patient.city}, {patient.state} {display_postal}"
            story.append(Paragraph(city_state_zip, delivery_address_style))
            
            # Build the document
            doc.build(story)
            self.logger.info(f"Generated Com-10 envelope PDF: {output_path}")
            
        except Exception as e:
            self.logger.error(f"Error generating envelope PDF: {e}")
            raise
    
    def _generate_csv_export(self, lines: List[InvoiceLine], output_path: Path):
        """Generate CSV line items export"""
        try:
            csv_data = []
            
            for line in lines:
                if line.is_previous_balance:
                    csv_data.append({
                        'Service Date': '',
                        'Description': line.description,
                        'Copay/Deductible': f"{line.amount:.2f}",
                        'Amount': '0'
                    })
                else:
                    csv_data.append({
                        'Service Date': line.service_date,
                        'Description': line.description,
                        'Copay/Deductible': '',
                        'Amount': f"{line.amount:.2f}"
                    })
            
            df = pd.DataFrame(csv_data)
            df.to_csv(output_path, index=False)
            self.logger.info(f"Generated CSV export: {output_path}")
            
        except Exception as e:
            self.logger.error(f"Error generating CSV export: {e}")
            raise
    
    def _generate_summary_report(self, summary: ProcessingSummary, output_dir: Path):
        """Generate comprehensive summary report"""
        try:
            summary_path = output_dir / f"Processing_Summary_{self.statement_date.strftime('%Y%m%d')}.txt"
            
            with open(summary_path, 'w') as f:
                f.write("=" * 80 + "\n")
                f.write("PATIENT INVOICE PROCESSING SUMMARY\n")
                f.write("=" * 80 + "\n\n")
                
                f.write(f"Processing Date: {summary.processing_date}\n")
                f.write(f"Statement Date: {self.statement_date.strftime('%B %d, %Y')}\n")
                f.write(f"Payment Due Date: {self.payment_due_date.strftime('%B %d, %Y')}\n")
                f.write(f"Amount Due Strategy: {self.amount_due_strategy}\n\n")
                
                f.write("SUMMARY STATISTICS:\n")
                f.write("-" * 40 + "\n")
                f.write(f"Total Patients Processed: {summary.total_processed}\n")
                f.write(f"Total Patients Skipped: {summary.total_skipped}\n")
                f.write(f"Total Errors: {summary.total_errors}\n")
                f.write(f"Total Amount Due: ${summary.total_amount_due:.2f}\n\n")
                
                if summary.processed_patients:
                    f.write("SUCCESSFULLY PROCESSED PATIENTS:\n")
                    f.write("-" * 40 + "\n")
                    for i, patient in enumerate(summary.processed_patients, 1):
                        f.write(f"{i:3d}. {patient}\n")
                    f.write("\n")
                
                if summary.skipped_patients:
                    f.write("SKIPPED PATIENTS:\n")
                    f.write("-" * 40 + "\n")
                    for i, (patient, reason) in enumerate(summary.skipped_patients, 1):
                        f.write(f"{i:3d}. {patient} - {reason}\n")
                    f.write("\n")
                
                if summary.errors:
                    f.write("ERRORS ENCOUNTERED:\n")
                    f.write("-" * 40 + "\n")
                    for i, (patient, error) in enumerate(summary.errors, 1):
                        f.write(f"{i:3d}. {patient} - {error}\n")
                    f.write("\n")
                
                f.write("FILES GENERATED:\n")
                f.write("-" * 40 + "\n")
                f.write("For each processed patient:\n")
                f.write("  - PDF Invoice: LastName_Year_Invoice_mmddyyyy.pdf\n")
                f.write("  - Cover Letter: LastName_Year_Envelope_mmddyyyy.docx\n")
                f.write("  - Envelope PDF: LastName_Year_Envelope_mmddyyyy.pdf\n")
                f.write("  - CSV Items: LastName_Year_InvoiceItems_mmddyyyy.csv\n\n")
                
                f.write("=" * 80 + "\n")
                f.write("End of Summary Report\n")
                f.write("=" * 80 + "\n")
            
            self.logger.info(f"Generated summary report: {summary_path}")
            
        except Exception as e:
            self.logger.error(f"Error generating summary report: {e}")
            raise
    
    def generate_invoices(self, roster_file: str, invoice_file: str, 
                         template_file: str, output_dir: str = "output",
                         custom_mapping: Dict = None, generate_csv: bool = True,
                         envelope_format: str = "both"):
        """Main method to generate all invoices and cover letters"""
        # Initialize summary tracking
        summary = ProcessingSummary(
            processed_patients=[],
            skipped_patients=[],
            errors=[],
            total_processed=0,
            total_skipped=0,
            total_errors=0,
            total_amount_due=0.0,
            processing_date=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )
        
        try:
            # Load data
            patients = self.load_patient_roster(roster_file)
            invoice_df = self.load_invoice_data(invoice_file, custom_mapping)
            
            # Create output directory
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)
            
            # Group by patient name
            patient_groups = invoice_df.groupby('name')
            
            for patient_name, patient_df in patient_groups:
                try:
                    # Match patient
                    patient, is_ambiguous = self._match_patient(patient_name, patients)
                    
                    if is_ambiguous:
                        self.logger.warning(f"Ambiguous match for patient: {patient_name}")
                    
                    # Generate invoice lines
                    lines, total_due, original_df = self._generate_invoice_lines(patient_df)
                    
                    if total_due <= 0:
                        reason = "No open balance"
                        self.logger.info(f"{reason} for {patient_name}, skipping")
                        summary.skipped_patients.append((patient_name, reason))
                        summary.total_skipped += 1
                        continue
                    
                    # Create patient folder and use matched patient data for naming
                    if patient:
                        # Use the MATCHED patient's name for folder and display, not the invoice name
                        folder_name = f"{patient.last_name}_{patient.first_name}_{patient.prn}"
                        patient_display_name = f"{patient.first_name} {patient.last_name} (PRN: {patient.prn})"
                        # Use the matched patient's data for all file generation
                        file_patient = patient
                    else:
                        first_name, last_name = self._parse_patient_name(patient_name)
                        folder_name = f"{last_name}_{first_name}_UNKNOWN"
                        patient_display_name = f"{first_name} {last_name} (No roster match)"
                        # Create dummy patient for file generation
                        file_patient = PatientData("", first_name, last_name, "", "", "", "", "", "")
                    
                    folder_name = self._sanitize_filename(folder_name)
                    patient_dir = output_path / folder_name
                    patient_dir.mkdir(exist_ok=True)
                    
                    # Generate filenames using the matched patient's last name
                    year = self.statement_date.year
                    date_str = self.statement_date.strftime("%m%d%Y")
                    base_name = f"{self._sanitize_filename(file_patient.last_name)}_{year}"
                    
                    # Generate files using the matched patient data
                    pdf_path = patient_dir / f"{base_name}_Invoice_{date_str}.pdf"
                    self._generate_pdf_invoice(file_patient, lines, total_due, original_df, pdf_path)

                    # Generate envelope in requested format(s)
                    if envelope_format in ["docx", "both"]:
                        docx_path = patient_dir / f"{base_name}_Envelope_{date_str}.docx"
                        self._generate_cover_letter(file_patient, template_file, docx_path)
                    
                    if envelope_format in ["pdf", "both"]:
                        envelope_pdf_path = patient_dir / f"{base_name}_Envelope_{date_str}.pdf"
                        self._generate_envelope_pdf(file_patient, template_file, envelope_pdf_path)
                    
                    if generate_csv:
                        csv_path = patient_dir / f"{base_name}_InvoiceItems_{date_str}.csv"
                        self._generate_csv_export(lines, csv_path)
                    
                    # Update summary
                    total_payments = float(original_df['paid'].sum())
                    summary.processed_patients.append(f"{patient_display_name} - Due: ${total_due:.2f}, Paid: ${total_payments:.2f}")
                    summary.total_processed += 1
                    summary.total_amount_due += total_due
                    
                except Exception as e:
                    error_msg = str(e)
                    self.logger.error(f"Error processing patient {patient_name}: {error_msg}")
                    summary.errors.append((patient_name, error_msg))
                    summary.total_errors += 1
                    continue
            
            # Generate summary report
            self._generate_summary_report(summary, output_path)
            
            self.logger.info(f"Processing complete. Generated: {summary.total_processed}, "
                           f"Skipped: {summary.total_skipped}, Errors: {summary.total_errors}")
            
            return summary
            
        except Exception as e:
            self.logger.error(f"Fatal error in invoice generation: {e}")
            # Still return summary even if there's a fatal error
            return summary


def main():
    """Example usage"""
    generator = PatientInvoiceGenerator(
        amount_due_strategy="auto"
        # Use today's date automatically
    )
    
    try:
        summary = generator.generate_invoices(
            roster_file="PatientListReport_active_20250912.csv",
            invoice_file="invoice_template.xlsx", 
            template_file="Access Multi Letter Cover.docx",
            output_dir="output",
            generate_csv=True,
            envelope_format="both"  # Options: "docx", "pdf", or "both"
        )
        
        if summary:
            print("Invoice generation completed successfully!")
            print(f"Processed: {summary.total_processed} patients")
            print(f"Skipped: {summary.total_skipped} patients")
            print(f"Errors: {summary.total_errors} patients")
            print(f"Total Amount Due: ${summary.total_amount_due:.2f}")
            print(f"Summary report saved in output directory")
        else:
            print("Invoice generation completed but no summary available")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
