"""
Automated Cold Email Outreach System
Sends personalized follow-up emails based on timing and response tracking
With rate limiting for Gemini API (1 min pause per 50 emails)
Auto-cleans CSV: fixes headings and converts full names to first names
"""

import os
import json
import base64
import pandas as pd
import time
import uuid
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import make_msgid
import random

from langchain_google_community import GmailToolkit
from langchain_google_community.gmail.utils import build_gmail_service, get_google_credentials

# Load environment
load_dotenv()

# Rate limiting configuration (Gemini API limits)
RATE_LIMIT_EMAILS = 50  # Process 50 emails
RATE_LIMIT_WAIT = 60    # Wait 60 seconds


def auto_clean_csv(csv_file: str) -> bool:
    """
    Automatically clean CSV: fix headings and extract first names
    Returns True if CSV is ready, False if there's an error
    """
    try:
        # Try multiple encodings to handle different file formats
        encodings = ['utf-8', 'latin-1', 'windows-1252', 'iso-8859-1', 'cp1252']
        df = None
        used_encoding = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv(csv_file, encoding=encoding)
                used_encoding = encoding
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        if df is None:
            print(f"❌ Could not read CSV with any supported encoding")
            return False
        
        if used_encoding != 'utf-8':
            print(f"  ℹ️ Detected encoding: {used_encoding} (will convert to UTF-8)")
        
        df.columns = df.columns.str.strip()
        
        # Fix duplicate column names (e.g., 'first_name,first_name' becomes 'first_name,contact_name')
        cols = df.columns.tolist()
        seen = {}
        for i, col in enumerate(cols):
            if col in seen:
                # Handle duplicate 'first_name' columns - assume second one is the actual name
                if col == 'first_name':
                    cols[i] = 'contact_name'
                else:
                    j = 2
                    while f"{col}_{j}" in seen:
                        j += 1
                    cols[i] = f"{col}_{j}"
            seen[col] = True
        df.columns = cols
        
        # Check if already in correct format
        has_correct_format = all(col in df.columns for col in ['company_name', 'first_name', 'email'])
        
        if has_correct_format:
            # Check if names look like first names only (no spaces = already first names)
            sample_names = df['first_name'].dropna().head(5)
            mostly_first_names = sum(' ' not in str(name) for name in sample_names) >= len(sample_names) * 0.6
            
            if mostly_first_names:
                print("✅ CSV already in correct format!")
                return True
        
        print("🔧 Auto-cleaning CSV...")
        
        # Smart column mapping
        column_mapping = {}
        for col in df.columns:
            col_lower = col.lower()
            if 'company' in col_lower or 'organization' in col_lower:
                column_mapping[col] = 'company_name'
            elif 'email' in col_lower or 'mail' in col_lower:
                column_mapping[col] = 'email'
            elif col == 'contact_name':  # This contains the actual full names
                column_mapping[col] = 'first_name'
            elif 'name' in col_lower and 'company' not in col_lower and 'email' not in col_lower and col != 'first_name':
                # Don't map the empty first_name column, use contact_name instead
                if 'contact_name' not in df.columns:
                    column_mapping[col] = 'first_name'
        
        df = df.rename(columns=column_mapping)
        
        # Clean up duplicate columns and ensure we keep only needed ones
        final_df = pd.DataFrame()
        
        # Get the best data for each required column
        if 'company_name' in df.columns:
            final_df['company_name'] = df['company_name']
        
        if 'email' in df.columns:
            final_df['email'] = df['email']
        
        # For first_name, get data from the column that actually has names
        first_name_data = None
        if 'first_name' in df.columns:
            # Check all possible name columns and pick the one with actual data
            for col in df.columns:
                if 'first_name' in col or 'contact_name' in col or col == 'first_name':
                    sample = df[col].dropna().head(3)
                    if len(sample) > 0 and not sample.iloc[0] == '':
                        first_name_data = df[col]
                        break
        
        if first_name_data is not None:
            final_df['first_name'] = first_name_data
        
        df = final_df
        
        # Extract first names from full names
        if 'first_name' in df.columns:
            def extract_first_name(full_name):
                # Handle NaN values properly
                try:
                    if pd.isna(full_name) or full_name is None:
                        return ""
                except (ValueError, TypeError):
                    # If pd.isna fails, treat as non-NaN
                    pass
                
                name_str = str(full_name).strip()
                if not name_str or name_str.lower() in ['nan', 'none', '']:
                    return ""
                
                parts = name_str.split()
                if parts:
                    first = parts[0]
                    # Remove titles
                    if first in ['Mr.', 'Mrs.', 'Ms.', 'Dr.', 'Prof.'] and len(parts) > 1:
                        first = parts[1]
                    return first.strip('.,')
                return name_str
            
            # Check if we need to extract first names (if they look like full names)
            sample_names = df['first_name'].dropna().head(5)
            needs_extraction = any(' ' in str(name) for name in sample_names)
            
            if needs_extraction:
                # Fix pandas FutureWarning by using proper apply function
                df['first_name'] = df['first_name'].apply(extract_first_name)
                print(f"  ✓ Converted full names to first names")
            else:
                print(f"  ✓ First names already extracted")
        
        # Validate required columns
        required = ['company_name', 'first_name', 'email']
        if not all(col in df.columns for col in required):
            print(f"❌ Could not find required columns. Found: {list(df.columns)}")
            return False
        
        # Save cleaned CSV (always in UTF-8)
        try:
            df[required].dropna(subset=['email']).to_csv(csv_file, index=False, encoding='utf-8')
            print(f"  ✓ Saved cleaned CSV with {len(df)} rows")
            print(f"  ✓ Columns: {required}")
            if used_encoding != 'utf-8':
                print(f"  ✓ Converted from {used_encoding} to UTF-8")
        except PermissionError:
            print(f"  ⚠️ Cannot save {csv_file} - file is open in another program")
            print(f"  ⚠️ Please close Excel/editor and run again, or rename the file")
            print(f"  ⚠️ Using original file format for now...")
            # Continue without saving - the cleaning will happen in memory
        
        return True
        
    except Exception as e:
        print(f"❌ Error cleaning CSV: {e}")
        import traceback
        traceback.print_exc()
        return False


# Email templates
EMAIL_TEMPLATES = {
    1: {
        "subject": "Remote Internship / Full-time",
        "body": """Hi {first_name},

I'm Gaurav Raj, a final-year student at IIT Roorkee. I have been following {company_name}'s work, and I see a strong alignment between your engineering goals and my background in scientific automation.
I am writing to apply for a Remote Internship or Full-time position where I can contribute immediately to your data or engineering team.

My Core Qualifications:
Scientific Python: At NIMS (Japan), I built a Python pipeline to automate TEM image analysis (FFT/Lorentzian fitting), reducing manual processing time by ~90%.
Production AI Agents: At GreenIntel, I worked on autonomous agents to handle complex user queries and API routing.

I have attached my resume and would welcome a brief 10-minute chat to discuss how I can contribute.

Best regards,
Gaurav Raj
<a href="https://www.linkedin.com/in/gaurav-raj-3ba84b254/">LinkedIn</a> | <a href="https://drive.google.com/file/d/1i7RHTFq9IHaPxQHrY_0jrowvVAkgo0Ot/view?usp=sharing">Resume Link</a>"""
    },
    2: {
        "subject": None,  # Reply to same thread
        "body": """Hi {first_name},
Just checking in to see if you had a chance to review this. Is there anything you need from me to move forward?

If you're not the right contact, could you please point me in the right direction?

Regards,
Gaurav Raj"""
    },
    3: {
        "subject": None,  # Reply to same thread
        "body": """Hello {first_name},

I just wanted to send one last note in case my previous emails got buried. I'd be truly grateful if you could respond to this email, even a quick "yes" or "no" would mean a lot.

Thanks so much for your time and consideration.

Best regards,
Gaurav Raj"""
    }
}

# Timing configuration (in days)
TIMING_CONFIG = {
    1: 0,   # Email 1: Send immediately
    2: 2,   # Email 2: After 2 days
    3: 5    # Email 3: After 5 days (2+3)
}


class ColdEmailAutomation:
    """Automated cold email outreach with follow-ups"""
    
    def __init__(self, excel_file: str, resume_path: str = "Gaurav_Resume.pdf", is_test_mode: bool = False):
        """
        Initialize the automation system
        
        Args:
            excel_file: Path to Excel/CSV with company data
            resume_path: Path to resume PDF file
            is_test_mode: If True, uses email_tracking_test.json
        """
        # Gmail API setup
        self.credentials = get_google_credentials(
            token_file="token.json",
            scopes=["https://mail.google.com/"],
            client_secrets_file="credentials.json",
        )
        self.service = build_gmail_service(credentials=self.credentials)
        
        # Get authenticated email address for per-account tracking
        try:
            profile = self.service.users().getProfile(userId='me').execute()
            self.authenticated_email = profile.get('emailAddress', 'default')
        except Exception as e:
            print(f"⚠️ Could not get authenticated email: {e}")
            self.authenticated_email = 'default'
        
        # Sanitize email for use in filename (replace @ and . with _)
        email_safe = self.authenticated_email.replace('@', '_at_').replace('.', '_')
        
        # Create tracking folder if it doesn't exist (keeps JSON files organized)
        self.tracking_folder = "tracking"
        os.makedirs(self.tracking_folder, exist_ok=True)
        
        # Load prospect data
        self.excel_file = excel_file
        self.resume_path = resume_path
        self.prospects = self.load_prospects()
        
        # Load/create tracking database (separate files for test vs production AND per email account)
        # All tracking files go in the tracking/ folder
        self.is_test_mode = is_test_mode
        if is_test_mode:
            self.tracking_file = os.path.join(self.tracking_folder, f"email_tracking_test_{email_safe}.json")
        else:
            self.tracking_file = os.path.join(self.tracking_folder, f"email_tracking_{email_safe}.json")
        
        # Archive file also per-account, in tracking folder
        self.archive_file = os.path.join(self.tracking_folder, f"email_archive_{email_safe}.json")
        
        # Archived domains file — persists domains where a reply was received
        self.archived_domains_file = os.path.join(self.tracking_folder, f"archived_domains_{email_safe}.json")
        
        self.tracking_db = self.load_tracking_db()
        self.archived_domains = self._load_archived_domains()
        
        print(f"✓ Authenticated as: {self.authenticated_email}")
        print(f"✓ Loaded {len(self.prospects)} prospects")
        print(f"✓ Tracking file: {self.tracking_file}")
        print(f"✓ Archive file: {self.archive_file}")
        print(f"✓ Archived domains file: {self.archived_domains_file}")
        print(f"✓ Tracking {len(self.tracking_db)} existing email threads")
        if self.archived_domains:
            print(f"✓ {len(self.archived_domains)} domain(s) already archived (replied): {sorted(self.archived_domains)}")
    
    def load_prospects(self) -> pd.DataFrame:
        """Load prospect data from Excel/CSV"""
        # Support either a single CSV/Excel file or a directory containing multiple CSVs
        if os.path.isdir(self.excel_file):
            # Read all .csv files in the directory (sorted by filename)
            csv_files = sorted([f for f in os.listdir(self.excel_file) if f.lower().endswith('.csv')])
            if not csv_files:
                raise ValueError(f"No CSV files found in directory: {self.excel_file}")

            frames = []
            for fname in csv_files:
                path = os.path.join(self.excel_file, fname)
                try:
                    df_part = pd.read_csv(path, encoding='utf-8')
                    # Add source file column to track which CSV each row came from
                    df_part['_source_csv'] = fname
                    frames.append(df_part)
                except UnicodeDecodeError:
                    try:
                        # Fallback to latin-1 encoding
                        df_part = pd.read_csv(path, encoding='latin-1')
                        df_part['_source_csv'] = fname
                        frames.append(df_part)
                        print(f"  ℹ️ {fname} read with latin-1 encoding")
                    except Exception as e:
                        print(f"⚠️ Warning: Failed to read {path}: {e}")
                except Exception as e:
                    print(f"⚠️ Warning: Failed to read {path}: {e}")

            if not frames:
                raise ValueError(f"No readable CSV files in directory: {self.excel_file}")

            df = pd.concat(frames, ignore_index=True)
            # Drop exact duplicate emails across files (keep first occurrence)
            if 'email' in df.columns:
                before = len(df)
                df = df.drop_duplicates(subset=['email'], keep='first').reset_index(drop=True)
                after = len(df)
                removed = before - after
                if removed > 0:
                    print(f"  ⚠️ Removed {removed} duplicate row(s) across CSV files based on email")
        else:
            if self.excel_file.endswith('.csv'):
                try:
                    df = pd.read_csv(self.excel_file, encoding='utf-8')
                except UnicodeDecodeError:
                    df = pd.read_csv(self.excel_file, encoding='latin-1')
                    print(f"  ℹ️ CSV read with latin-1 encoding")
                # Add source file for single CSV
                df['_source_csv'] = os.path.basename(self.excel_file)
            else:
                df = pd.read_excel(self.excel_file)
                df['_source_csv'] = os.path.basename(self.excel_file)
        
        # Normalize column names to match expected format
        # Support: (company_name, first_name, email) OR (Company, Name, Email)
        df.columns = df.columns.str.strip()

        # Create mapping for different column name variations
        column_mapping = {}
        for col in df.columns:
            col_lower = col.lower()
            if 'company' in col_lower or 'organization' in col_lower:
                column_mapping[col] = 'company_name'
            elif 'name' in col_lower and 'company' not in col_lower:
                column_mapping[col] = 'first_name'
            elif 'email' in col_lower or 'mail' in col_lower:
                column_mapping[col] = 'email'

        # Rename columns
        df = df.rename(columns=column_mapping)

        # Verify required columns exist
        required_cols = ['company_name', 'first_name', 'email']
        for col in required_cols:
            if col not in df.columns:
                raise ValueError(f"Missing required column: {col}. Found columns: {list(df.columns)}")

        # Expand rows that contain multiple emails in one cell (split by newline, semicolon, comma)
        def split_multi_emails(row):
            emails = str(row['email']).replace(';', '\n').replace(',', '\n').split('\n')
            cleaned = [e.strip() for e in emails if e and str(e).strip()]
            rows = []
            for e in cleaned:
                new_row = {
                    'company_name': str(row.get('company_name', '')).strip(),
                    'first_name': str(row.get('first_name', '')).strip(),
                    'email': e.lower(),
                    '_source_csv': str(row.get('_source_csv', '')).strip()
                }
                rows.append(new_row)
            return rows

        expanded = []
        for _, r in df.iterrows():
            try:
                parts = split_multi_emails(r)
                expanded.extend(parts)
            except Exception:
                # Fallback: include the raw row
                expanded.append({
                    'company_name': str(r.get('company_name', '')).strip(),
                    'first_name': str(r.get('first_name', '')).strip(),
                    'email': str(r.get('email', '')).strip().lower(),
                    '_source_csv': str(r.get('_source_csv', '')).strip()
                })

        new_df = pd.DataFrame(expanded)

        # Validate email addresses with a simple regex and drop invalid ones
        import re
        email_regex = re.compile(r"^[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,}$")

        def valid_email(e):
            try:
                return bool(email_regex.match(str(e).strip()))
            except Exception:
                return False

        before = len(new_df)
        new_df = new_df[new_df['email'].apply(valid_email)].reset_index(drop=True)
        after = len(new_df)
        removed = before - after
        if removed > 0:
            print(f"  ⚠️ Removed {removed} invalid email row(s) during load_prospects")

        # Ensure columns order
        return new_df[['company_name', 'first_name', 'email', '_source_csv']]
    
    def load_tracking_db(self) -> Dict:
        """
        Load email tracking database with automatic recovery
        Handles empty/corrupted JSON files automatically
        """
        if os.path.exists(self.tracking_file):
            try:
                with open(self.tracking_file, 'r') as f:
                    content = f.read().strip()
                    
                    # Check if file is empty
                    if not content:
                        print(f"⚠️ Warning: {self.tracking_file} is empty!")
                        print(f"🔧 Auto-fixing: Initializing with empty JSON...")
                        # Auto-fix: Write empty JSON
                        with open(self.tracking_file, 'w') as fix_f:
                            json.dump({}, fix_f, indent=2)
                        print(f"✅ Fixed: {self.tracking_file} initialized")
                        print(f"💡 Smart Recovery will rebuild history from Gmail as needed\n")
                        return {}
                    
                    # Try to parse JSON
                    try:
                        return json.loads(content)
                    except json.JSONDecodeError as je:
                        print(f"⚠️ Warning: {self.tracking_file} has invalid JSON!")
                        print(f"   Error: {str(je)}")
                        print(f"🔧 Auto-fixing: Creating backup and initializing fresh...")
                        
                        # Create backup of corrupted file
                        backup_file = self.tracking_file.replace('.json', '_corrupted_backup.json')
                        with open(backup_file, 'w') as backup_f:
                            backup_f.write(content)
                        print(f"💾 Backup saved: {backup_file}")
                        
                        # Write fresh empty JSON
                        with open(self.tracking_file, 'w') as fix_f:
                            json.dump({}, fix_f, indent=2)
                        print(f"✅ Fixed: {self.tracking_file} initialized")
                        print(f"💡 Smart Recovery will rebuild history from Gmail as needed\n")
                        return {}
                        
            except Exception as e:
                print(f"⚠️ Error reading {self.tracking_file}: {e}")
                print(f"🔧 Creating fresh tracking file...")
                with open(self.tracking_file, 'w') as f:
                    json.dump({}, f, indent=2)
                print(f"✅ Fixed: {self.tracking_file} initialized\n")
                return {}
        
        # File doesn't exist - create it
        print(f"📁 Creating new tracking file: {self.tracking_file}")
        with open(self.tracking_file, 'w') as f:
            json.dump({}, f, indent=2)
        print(f"✅ Created: {self.tracking_file}\n")
        return {}
    
    def save_tracking_db(self):
        """Save email tracking database with error handling"""
        try:
            # Atomic write: write to temp file then replace
            tmp_file = f"{self.tracking_file}.tmp"
            with open(tmp_file, 'w') as f:
                json.dump(self.tracking_db, f, indent=2)
            os.replace(tmp_file, self.tracking_file)
        except IOError as e:
            print(f"⚠️ Error saving tracking database: {e}")
            print(f"⚠️ This could be due to disk full or permission issues")
            print(f"⚠️ Your progress may not be saved!")

    # --- Archive helpers -------------------------------------------------

    # Generic/personal email domains — we do NOT batch-archive these because
    # different people at gmail.com / yahoo.com are unrelated companies.
    GENERIC_DOMAINS = {
        'gmail.com', 'yahoo.com', 'yahoo.in', 'hotmail.com', 'outlook.com',
        'live.com', 'icloud.com', 'me.com', 'aol.com', 'protonmail.com',
        'proton.me', 'zoho.com', 'mail.com', 'ymail.com', 'googlemail.com',
    }

    def _get_domain(self, email: str) -> str:
        """Extract the domain part of an email address.
        
        Returns empty string if email is invalid.
        """
        try:
            return email.strip().lower().split('@')[1]
        except (IndexError, AttributeError):
            return ''

    def _load_archived_domains(self) -> set:
        """Load the persisted set of archived domains from disk."""
        if os.path.exists(self.archived_domains_file):
            try:
                with open(self.archived_domains_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # stored as list of objects: [{"domain": ..., "archived_at": ..., "reason": ...}]
                    return set(entry['domain'] for entry in data if 'domain' in entry)
            except Exception:
                return set()
        return set()

    def _save_archived_domains(self, domain: str, reason: str = 'replied'):
        """Append a domain to the persisted archived_domains JSON file."""
        try:
            existing: list = []
            if os.path.exists(self.archived_domains_file):
                try:
                    with open(self.archived_domains_file, 'r', encoding='utf-8') as f:
                        existing = json.load(f)
                except Exception:
                    existing = []
            
            # Avoid duplicate entries
            known = {e['domain'] for e in existing if 'domain' in e}
            if domain not in known:
                existing.append({
                    'domain': domain,
                    'archived_at': datetime.now().isoformat(),
                    'reason': reason,
                })
            
            tmp = self.archived_domains_file + '.tmp'
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(existing, f, indent=2, ensure_ascii=False)
            os.replace(tmp, self.archived_domains_file)
        except Exception as e:
            print(f"  ⚠️ Could not save archived domains: {e}")

    def _load_archive(self) -> dict:
        """Load archive file if exists, otherwise return empty dict."""
        archive_file = getattr(self, 'archive_file', 'email_archive.json')
        if os.path.exists(archive_file):
            try:
                with open(archive_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _save_archive(self, archive_data: dict):
        """Save archive data atomically to archive file."""
        archive_file = getattr(self, 'archive_file', 'email_archive.json')
        tmp = archive_file + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(archive_data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, archive_file)

    def archive_domain(self, replied_email: str, reason: str = 'replied') -> int:
        """Archive ALL contacts that share the same email domain as `replied_email`.

        This prevents sending further emails to colleagues at the same company
        once any one person has responded.

        Generic personal domains (gmail.com, yahoo.com, etc.) are intentionally
        skipped — we only archive under a domain when it clearly belongs to a
        single organisation.

        Args:
            replied_email: The email address that sent a reply.
            reason: Archive reason tag stored in the record.

        Returns:
            Number of additional prospects archived (excluding the replier).
        """
        domain = self._get_domain(replied_email)
        if not domain:
            print(f"  ⚠️ Could not extract domain from {replied_email}; skipping domain archive.")
            return 0

        if domain in self.GENERIC_DOMAINS:
            print(f"  ℹ️ Domain '{domain}' is a generic/personal provider — skipping batch archive.")
            return 0

        if domain in self.archived_domains:
            print(f"  ℹ️ Domain '{domain}' already archived — nothing to do.")
            return 0

        print(f"  🏢 Archiving all contacts from domain: @{domain}")

        # Collect keys to archive (copy to avoid mutating dict while iterating)
        to_archive = [
            key for key, record in list(self.tracking_db.items())
            if self._get_domain(key) == domain and key != replied_email
        ]

        archived_count = 0
        for key in to_archive:
            try:
                company = self.tracking_db.get(key, {}).get('company_name', key)
                name = self.tracking_db.get(key, {}).get('first_name', '')
                print(f"    ↳ Archiving colleague: {name} ({key}) at {company}")
                self.archive_prospect(key, reason=reason)
                archived_count += 1
            except Exception as e:
                print(f"    ⚠️ Failed to archive {key}: {e}")

        # Mark domain as archived (in-memory + persistent)
        self.archived_domains.add(domain)
        self._save_archived_domains(domain, reason=reason)

        if archived_count:
            print(f"  ✅ Domain archive complete: {archived_count} additional contact(s) archived from @{domain}")
        else:
            print(f"  ✅ Domain archive complete: no other active contacts found for @{domain}")

        return archived_count

    def archive_prospect(self, tracking_key: str, reason: str = "completed") -> bool:
        """Move a prospect from tracking_db into the archive file and remove from CSV.
        Also saves to completed/ folder in appropriate CSV file.

        Args:
            tracking_key: Email address of prospect
            reason: "completed" for 3 emails sent or reply received, "bounced" for invalid email
            
        Returns True if archived successfully, False otherwise.
        """
        if tracking_key not in self.tracking_db:
            return False

        try:
            archive = self._load_archive()
            record = self.tracking_db.pop(tracking_key)
            record['_archived_at'] = datetime.now().isoformat()
            record['_archive_reason'] = reason
            
            # Get source CSV file to remove from
            source_csv = record.get('_source_csv', '')

            # If key exists, convert to list or append
            if tracking_key in archive:
                existing = archive[tracking_key]
                if isinstance(existing, list):
                    existing.append(record)
                    archive[tracking_key] = existing
                else:
                    archive[tracking_key] = [existing, record]
            else:
                archive[tracking_key] = record

            self._save_archive(archive)
            # Persist tracking DB removal
            self.save_tracking_db()
            
            # Save to completed/ folder CSV files
            self._save_to_completed_csv(record, reason)
            
            # Remove from source CSV file
            if source_csv:
                self._remove_from_csv(tracking_key, source_csv)
            else:
                # FALLBACK: If source_csv is missing, try to remove from all CSV files
                print(f"  ⚠️ No source CSV recorded - searching all CSV files...")
                self._remove_from_all_csvs(tracking_key)
            
            print(f"  📦 Archived prospect {tracking_key} ({reason}) and removed from CSV")
            return True
        except Exception as e:
            print(f"  ❌ Failed to archive {tracking_key}: {e}")
            return False
            
    def _save_to_completed_csv(self, record: Dict, reason: str):
        """Save archived prospect to completed/ folder CSV files."""
        try:
            # Create completed/ folder if it doesn't exist
            completed_dir = "completed"
            os.makedirs(completed_dir, exist_ok=True)
            
            # Determine which CSV file to use
            if reason == "bounced":
                csv_filename = os.path.join(completed_dir, "bounced_emails.csv")
            else:
                csv_filename = os.path.join(completed_dir, "completed_prospects.csv")
            
            # Prepare row data
            row_data = {
                'company_name': record.get('company_name', ''),
                'first_name': record.get('first_name', ''),
                'email': record.get('email', ''),
                'emails_sent': len(record.get('emails_sent', [])),
                'received_reply': record.get('received_reply', False),
                'reply_at_stage': record.get('reply_at_stage', ''),
                'archived_at': record.get('_archived_at', ''),
                'archive_reason': reason
            }
            
            # Check if CSV exists
            if os.path.exists(csv_filename):
                # Append to existing CSV
                df_existing = pd.read_csv(csv_filename)
                df_new = pd.DataFrame([row_data])
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                df_combined.to_csv(csv_filename, index=False, encoding='utf-8')
            else:
                # Create new CSV
                df = pd.DataFrame([row_data])
                df.to_csv(csv_filename, index=False, encoding='utf-8')
            
            print(f"  💾 Saved to {csv_filename}")
            
        except Exception as e:
            print(f"  ⚠️ Failed to save to completed CSV: {e}")
    
    def _remove_from_csv(self, email_to_remove: str, source_csv: str):
        """Remove a specific email from the source CSV file."""
        try:
            # Determine full path to CSV
            if os.path.isdir(self.excel_file):
                csv_path = os.path.join(self.excel_file, source_csv)
            else:
                csv_path = self.excel_file
            
            if not os.path.exists(csv_path):
                print(f"  ⚠️ CSV file not found: {csv_path}")
                return
            
            # Read CSV
            try:
                df = pd.read_csv(csv_path, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(csv_path, encoding='latin-1')
            
            # Remove the row with matching email
            original_count = len(df)
            
            # Handle different column name variations
            email_col = None
            for col in df.columns:
                if 'email' in col.lower() or 'mail' in col.lower():
                    email_col = col
                    break
            
            if email_col is None:
                print(f"  ⚠️ No email column found in {source_csv}")
                return
            
            # Remove rows matching this email (case insensitive)
            # Handle both single emails and multi-line cells with multiple emails
            def contains_email(cell_value):
                if pd.isna(cell_value):
                    return False
                cell_lower = str(cell_value).lower()
                email_lower = email_to_remove.lower()
                # Check if email matches exactly OR is contained in multi-line cell
                return cell_lower == email_lower or email_lower in cell_lower
            
            df = df[~df[email_col].apply(contains_email)]
            new_count = len(df)
            
            if new_count < original_count:
                # Save back to CSV (preserve original encoding)
                df.to_csv(csv_path, index=False, encoding='utf-8')
                removed = original_count - new_count
                print(f"  ✂️ Removed {removed} row(s) from {source_csv}")
            else:
                print(f"  ℹ️ Email {email_to_remove} not found in {source_csv}")
                
        except Exception as e:
            print(f"  ⚠️ Failed to remove from CSV {source_csv}: {e}")
    
    def _remove_from_all_csvs(self, email_to_remove: str):
        """Remove a specific email from all CSV files in the directory.
        Used as fallback when source_csv is not known."""
        try:
            # Check if excel_file is a directory with multiple CSVs
            if os.path.isdir(self.excel_file):
                csv_files = [f for f in os.listdir(self.excel_file) if f.lower().endswith('.csv')]
                
                if not csv_files:
                    print(f"  ⚠️ No CSV files found in directory")
                    return
                
                total_removed = 0
                for csv_file in csv_files:
                    csv_path = os.path.join(self.excel_file, csv_file)
                    
                    try:
                        # Read CSV
                        try:
                            df = pd.read_csv(csv_path, encoding='utf-8')
                        except UnicodeDecodeError:
                            df = pd.read_csv(csv_path, encoding='latin-1')
                        
                        original_count = len(df)
                        
                        # Find email column
                        email_col = None
                        for col in df.columns:
                            if 'email' in col.lower() or 'mail' in col.lower():
                                email_col = col
                                break
                        
                        if email_col is None:
                            continue
                        
                        # Remove rows matching this email
                        df = df[df[email_col].str.lower() != email_to_remove.lower()]
                        new_count = len(df)
                        
                        if new_count < original_count:
                            # Save back to CSV
                            df.to_csv(csv_path, index=False, encoding='utf-8')
                            removed = original_count - new_count
                            total_removed += removed
                            print(f"  ✂️ Removed {removed} row(s) from {csv_file}")
                    
                    except Exception as e:
                        print(f"  ⚠️ Error processing {csv_file}: {e}")
                        continue
                
                if total_removed == 0:
                    print(f"  ℹ️ Email {email_to_remove} not found in any CSV file")
                else:
                    print(f"  ✅ Total removed: {total_removed} row(s) across all CSV files")
            else:
                # Single CSV file - just use _remove_from_csv
                source_csv = os.path.basename(self.excel_file)
                self._remove_from_csv(email_to_remove, source_csv)
                
        except Exception as e:
            print(f"  ⚠️ Failed to remove from CSV files: {e}")
    
    def reconstruct_tracking_from_gmail(self, email_address: str, company_name: str, first_name: str, source_csv: str = '') -> Optional[Dict]:
        """
        SMART RECOVERY: Reconstruct tracking data from Gmail sent folder
        This rebuilds the JSON tracking if it was deleted but emails were sent
        
        Args:
            email_address: Email to check
            company_name: Company name for the prospect
            first_name: First name for the prospect
            source_csv: Source CSV filename for this prospect
        
        Returns:
            Reconstructed tracking data dict, or None if no emails found
        """
        try:
            print(f"  🔍 Checking Gmail sent folder for existing emails...")
            
            # Search SENT emails TO this address (get ALL to rebuild full history)
            query = f"to:{email_address} in:sent"
            results = self.service.users().messages().list(
                userId='me',
                q=query,
                maxResults=10  # Get up to 10 emails (should cover all 3 stages)
            ).execute()
            
            messages = results.get('messages', [])
            
            if not messages:
                print(f"  ✓ No previous emails found in Gmail")
                return None
            
            print(f"  📧 Found {len(messages)} sent email(s) to {email_address}")
            print(f"  🔧 RECONSTRUCTING tracking data from Gmail...")
            
            # Reconstruct tracking data
            emails_sent = []
            
            for msg in messages:
                msg_id = msg['id']
                
                # Get email details (need thread_id, message_id, date, subject)
                msg_data = self.service.users().messages().get(
                    userId='me',
                    id=msg_id,
                    format='metadata',
                    metadataHeaders=['Subject', 'Date', 'Message-ID', 'Message-Id']
                ).execute()
                
                thread_id = msg_data.get('threadId')
                internal_date = int(msg_data.get('internalDate', 0))
                sent_date = datetime.fromtimestamp(internal_date / 1000).isoformat()
                
                # Extract Message-ID header
                headers = msg_data.get('payload', {}).get('headers', [])
                email_message_id = None
                subject = None
                
                for h in headers:
                    if h['name'] in ['Message-ID', 'Message-Id']:
                        # Normalize to ensure angle brackets
                        val = str(h['value']).strip()
                        if not (val.startswith('<') and val.endswith('>')) and '@' in val:
                            val = f"<{val}>"
                        email_message_id = val
                    elif h['name'] == 'Subject':
                        subject = h['value']
                
                # Determine stage based on subject
                stage = 1  # Default to stage 1
                if subject:
                    if subject.startswith("Re:"):
                        # Follow-up email
                        # Count how many follow-ups we have
                        followup_count = sum(1 for e in emails_sent if e['stage'] > 1)
                        stage = 2 + followup_count
                    else:
                        # First email (has subject with company name)
                        stage = 1
                
                emails_sent.append({
                    'stage': stage,
                    'message_id': msg_id,
                    'thread_id': thread_id,
                    'email_message_id': email_message_id or msg_id,
                    'sent_date': sent_date,
                    'reconstructed': True  # Flag to indicate this was rebuilt
                })
            
            # Sort by date (oldest first)
            emails_sent.sort(key=lambda x: x['sent_date'])
            
            # Reassign stages correctly (1, 2, 3) - limit to 3 emails max
            final_emails = []
            for idx, email in enumerate(emails_sent[:3], start=1):  # Only take first 3 emails
                email['stage'] = idx
                final_emails.append(email)
            
            emails_sent = final_emails  # Keep only the first 3 emails
            
            # Check for replies - use first email's thread for comprehensive check
            thread_id_to_check = emails_sent[0]['thread_id'] if emails_sent else None
            received_reply = False
            reply_at_stage = None
            
            if thread_id_to_check:
                print(f"  🔍 Checking if prospect replied in thread...")
                received_reply = self.check_for_reply(email_address, thread_id_to_check)
                
                # If reply found, determine at which stage
                if received_reply:
                    # They replied after the last email we sent
                    reply_at_stage = emails_sent[-1]['stage'] if emails_sent else 1
            
            tracking_data = {
                'company_name': company_name,
                'first_name': first_name,
                'email': email_address,
                'emails_sent': emails_sent,
                'received_reply': received_reply,
                'reply_at_stage': reply_at_stage,
                'reconstructed_from_gmail': True,
                'reconstruction_date': datetime.now().isoformat(),
                '_source_csv': source_csv  # CRITICAL: Store source CSV for later archiving
            }
            
            print(f"  ✅ RECONSTRUCTED: Found {len(emails_sent)} email(s) at stage(s): {[e['stage'] for e in emails_sent]}")
            if received_reply:
                print(f"  ✅ Reply detected from prospect after stage {reply_at_stage}")
            
            return tracking_data
            
        except Exception as e:
            print(f"  ⚠️ Error reconstructing from Gmail: {e}")
            return None
    
    def extract_email_from_header(self, from_header: str) -> str:
        """
        Extract email address from 'From' header value
        Handles formats like: 'Name <email@example.com>' or 'email@example.com'
        
        Args:
            from_header: The 'From' header value
        
        Returns:
            Extracted email address in lowercase
        """
        try:
            if not from_header:
                return ''
            
            # Check if format is 'Name <email@example.com>'
            if '<' in from_header and '>' in from_header:
                start = from_header.index('<') + 1
                end = from_header.index('>')
                return from_header[start:end].strip().lower()
            else:
                # Simple format: just email address
                return from_header.strip().lower()
        except:
            return from_header.strip().lower()
    
    def check_for_bounce(self, email_address: str, thread_id: Optional[str] = None) -> bool:
        """Check if email bounced (undeliverable, address not found, etc.)
        
        Args:
            email_address: Email to check
            thread_id: Thread ID to check for bounce notifications
            
        Returns:
            True if bounce detected, False otherwise
        """
        try:
            # Check for bounce notifications in thread
            if thread_id and thread_id.strip():
                try:
                    thread_data = self.service.users().threads().get(
                        userId='me',
                        id=thread_id
                    ).execute()
                    
                    messages = thread_data.get('messages', [])
                    
                    # Check messages for bounce indicators
                    for msg in messages:
                        headers = msg.get('payload', {}).get('headers', [])
                        
                        # Get sender
                        sender = ''
                        for h in headers:
                            if h.get('name') == 'From':
                                sender = h.get('value', '').lower()
                                break
                        
                        # Check if from mailer-daemon or similar
                        bounce_senders = [
                            'mailer-daemon@',
                            'postmaster@',
                            'mail delivery',
                            'delivery status'
                        ]
                        
                        if any(pattern in sender for pattern in bounce_senders):
                            # Get subject to confirm it's a bounce
                            subject = ''
                            for h in headers:
                                if h.get('name') == 'Subject':
                                    subject = h.get('value', '').lower()
                                    break
                            
                            bounce_keywords = [
                                'undeliverable',
                                'delivery failed',
                                'returned mail',
                                'address not found',
                                'user unknown',
                                'mailbox unavailable',
                                'does not exist',
                                'permanent error',
                                'failure notice',
                                'delivery status notification'
                            ]
                            
                            # Check subject for bounce keywords
                            if any(keyword in subject for keyword in bounce_keywords):
                                print(f"      ⚠️ BOUNCE DETECTED: {subject[:50]}...")
                                return True
                            
                            # Also check body snippet
                            snippet = msg.get('snippet', '').lower()
                            if any(keyword in snippet for keyword in bounce_keywords):
                                print(f"      ⚠️ BOUNCE DETECTED in message body")
                                return True
                            
                            # If from mailer-daemon but no keywords, still treat as bounce
                            # (mailer-daemon only sends bounce notifications)
                            if 'mailer-daemon@' in sender:
                                print(f"      ⚠️ BOUNCE DETECTED: mailer-daemon message found")
                                print(f"      📧 Subject: {subject[:80] if subject else 'No subject'}")
                                return True
                            
                except Exception as e:
                    print(f"  ⚠️ Error checking for bounce in thread: {e}")
            
            # Also check for bounce emails TO us about this address
            try:
                query = f'from:mailer-daemon OR from:postmaster subject:({email_address})'
                results = self.service.users().messages().list(
                    userId='me',
                    q=query,
                    maxResults=5
                ).execute()
                
                messages = results.get('messages', [])
                if messages:
                    print(f"      ⚠️ BOUNCE DETECTED: Found {len(messages)} delivery failure notification(s)")
                    return True
                    
            except Exception as e:
                print(f"  ⚠️ Error checking for bounce notifications: {e}")
            
            return False
            
        except Exception as e:
            print(f"  ⚠️ Error in bounce detection: {e}")
            return False
    
    def check_for_reply(self, email_address: str, thread_id: Optional[str] = None) -> bool:
        """
        Check if we received a reply in the email thread OR from the specific address
        COMPREHENSIVE FIX: Handles all edge cases including:
        - Replies from CC'd people in thread
        - Complex email header formats
        - API errors with proper fallbacks
        - Null/missing thread IDs
        - Race conditions
        
        Args:
            email_address: Email to check
            thread_id: Specific thread ID to check (optional)
        
        Returns:
            True if reply found (in thread OR as new email), False otherwise
        """
        max_retries = 3
        reply_found = False
        
        # PRIORITY 1: Check if thread has ANY replies (including from CC'd people)
        if thread_id and thread_id.strip():  # Null-safe check
            for attempt in range(max_retries):
                try:
                    # Get all messages in the thread
                    thread_data = self.service.users().threads().get(
                        userId='me',
                        id=thread_id
                    ).execute()
                    
                    messages = thread_data.get('messages', [])
                    
                    # We sent the first email, so check if there are more messages (replies)
                    if len(messages) > 1:
                        # Get our email address once
                        try:
                            profile = self.service.users().getProfile(userId='me').execute()
                            my_email = profile.get('emailAddress', '').lower()
                        except:
                            # Fallback: assume common Gmail pattern
                            my_email = 'unknown@gmail.com'
                        
                        # Check each message in thread (skip the first one which is ours)
                        for msg in messages[1:]:
                            try:
                                # Get sender from headers
                                headers = msg.get('payload', {}).get('headers', [])
                                sender_header = None
                                
                                for h in headers:
                                    if h.get('name') == 'From':
                                        sender_header = h.get('value', '')
                                        break
                                
                                if not sender_header:
                                    continue
                                
                                # Extract email from header (handles complex formats)
                                sender_email = self.extract_email_from_header(sender_header)
                                
                                # Filter out system/automated emails (bounce notifications, no-reply, etc.)
                                system_senders = [
                                    'mailer-daemon@',
                                    'no-reply@',
                                    'noreply@',
                                    'postmaster@',
                                    'bounces@',
                                    'bounce@',
                                    'do-not-reply@',
                                    'donotreply@'
                                ]
                                
                                is_system_email = any(pattern in sender_email.lower() for pattern in system_senders)
                                
                                # If sender is not us AND not a system email, it's a real reply
                                if sender_email and my_email not in sender_email and sender_email != my_email and not is_system_email:
                                    print(f"      → Reply found IN thread from: {sender_email}")
                                    reply_found = True
                                    return True
                                elif is_system_email:
                                    print(f"      ⚠️ Ignoring system email from: {sender_email}")
                            except Exception as msg_error:
                                # Skip this message if error, continue checking others
                                continue
                    
                    break  # Success, exit retry loop
                    
                except Exception as e:
                    if attempt < max_retries - 1:
                        print(f"  ⚠️ Retry {attempt + 1}/{max_retries} for thread check...")
                        time.sleep(2)
                        continue
                    else:
                        print(f"  ⚠️ Error checking thread for replies (will check direct emails): {e}")
                        # Don't return False yet, continue to check direct emails
        
        # PRIORITY 2: Check for direct emails FROM the specific address
        if not reply_found:
            for attempt in range(max_retries):
                try:
                    # Search for emails FROM this address (any conversation)
                    query = f"from:{email_address}"
                    results = self.service.users().messages().list(
                        userId='me',
                        q=query,
                        maxResults=10
                    ).execute()
                    
                    messages = results.get('messages', [])
                    
                    if not messages:
                        return False
                    
                    # Check if ANY email received from this address
                    if thread_id and thread_id.strip():
                        # Check if reply is in our thread specifically
                        in_thread = any(msg.get('threadId') == thread_id for msg in messages)
                        # Check if ANY email from them exists
                        any_reply = len(messages) > 0
                        
                        if in_thread:
                            print(f"      → Reply found from {email_address} IN thread")
                            return True
                        elif any_reply:
                            print(f"      → Reply found from {email_address} as NEW conversation")
                            return True
                    else:
                        # No thread ID provided, any email from this address counts as reply
                        if len(messages) > 0:
                            print(f"      → Reply found from {email_address}")
                            return True
                    
                    break  # Success, exit retry loop
                    
                except Exception as e:
                    if attempt < max_retries - 1:
                        print(f"  ⚠️ Retry {attempt + 1}/{max_retries} for direct email check...")
                        time.sleep(2)
                        continue
                    else:
                        print(f"  ⚠️ Error checking for direct reply: {e}")
                        return False
        
        return False
    
    def create_message_with_attachment(self, to: str, subject: str, body: str, 
                                      attachment_path: Optional[str] = None) -> Tuple[Dict, str]:
        """Create email message with optional attachment"""
        # Use 'mixed' for attachments, 'alternative' for text/html only
        message = MIMEMultipart('mixed')
        message['To'] = to
        message['Subject'] = subject
        # Generate a proper Message-ID for threading
        message['Message-ID'] = make_msgid()
        
        # Create alternative part for text and HTML
        msg_alternative = MIMEMultipart('alternative')
        
        # Add BOTH plain text and HTML versions
        # Plain text version (fallback)
        plain_text = MIMEText(body, 'plain')
        msg_alternative.attach(plain_text)
        
        # HTML version (with clickable links)
        html_body = body.replace('\n', '<br>')
        html_text = MIMEText(html_body, 'html')
        msg_alternative.attach(html_text)
        
        # Attach the alternative part to main message
        message.attach(msg_alternative)
        
        # Add attachment if provided
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as f:
                # Use proper MIME type for PDF files
                filename = os.path.basename(attachment_path)
                if filename.lower().endswith('.pdf'):
                    part = MIMEBase('application', 'pdf')
                else:
                    part = MIMEBase('application', 'octet-stream')
                
                part.set_payload(f.read())
                encoders.encode_base64(part)
                
                # Add proper headers for better compatibility
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{filename}"'
                )
                part.add_header(
                    'Content-Type',
                    f'application/pdf; name="{filename}"' if filename.lower().endswith('.pdf') else f'application/octet-stream; name="{filename}"'
                )
                
                message.attach(part)
            print(f"  📎 Attachment added: {os.path.basename(attachment_path)}")
        else:
            if attachment_path:
                print(f"  ⚠️ Warning: Attachment not found: {attachment_path}")
        
        # Store the Message-ID we generated
        msg_id = message['Message-ID']
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        return {'raw': raw}, msg_id  # Return both message and Message-ID
    
    def create_reply_message(self, to: str, body: str, thread_id: str, message_id: str, subject: Optional[str] = None) -> Dict:
        """Create a reply message in existing thread"""
        # Safety check
        if not body or not body.strip():
            raise ValueError(f"Email body cannot be empty! Received body: '{body}'")
        
        message = MIMEMultipart('alternative')  # Support both plain and HTML
        message['To'] = to
        message['Subject'] = subject if subject else "Re: (no subject)"

        # Ensure In-Reply-To and References use proper Message-ID format
        def normalize_msg_id(mid: str) -> str:
            if not mid:
                return ''
            mid = str(mid).strip()
            # If it already has angle brackets, return as-is
            if mid.startswith('<') and mid.endswith('>'):
                return mid
            # If it looks like a Message-ID (contains @), wrap in brackets
            if '@' in mid:
                return f"<{mid}>"
            return mid

        normalized_mid = normalize_msg_id(message_id)
        if normalized_mid:
            message['In-Reply-To'] = normalized_mid
            message['References'] = normalized_mid
        
        # Add BOTH plain text and HTML versions
        plain_text = MIMEText(body, 'plain')
        message.attach(plain_text)
        
        html_body = body.replace('\n', '<br>')
        html_text = MIMEText(html_body, 'html')
        message.attach(html_text)
        
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        return {'raw': raw, 'threadId': thread_id}
    
    def create_draft(self, message: Dict) -> Optional[str]:
        """Create a draft email"""
        try:
            draft = self.service.users().drafts().create(
                userId='me',
                body={'message': message}
            ).execute()
            
            return draft['id']
        except Exception as e:
            print(f"  ❌ Error creating draft: {e}")
            return None
    
    def send_message(self, message: Dict) -> Optional[str]:
        """Send an email message with retry logic"""
        max_retries = 5
        base_delay = 3  # seconds

        for attempt in range(max_retries):
            try:
                print(f"  📤 Sending email..." + (f" (attempt {attempt + 1}/{max_retries})" if attempt > 0 else ""))
                sent_message = self.service.users().messages().send(
                    userId='me',
                    body=message
                ).execute()

                message_id = sent_message.get('id')
                if message_id:
                    print(f"  ✅ Message sent! ID: {message_id}")
                return message_id

            except Exception as e:
                error_str = str(e).lower()
                # Handle common recoverable errors with backoff
                if 'rate limit' in error_str or '429' in error_str or 'quota' in error_str:
                    delay = base_delay * (attempt + 1)
                    print(f"  ⚠️ Rate limit hit. Waiting {delay} seconds...")
                    time.sleep(delay)
                    continue
                if error_str.startswith('5') or 'internalerror' in error_str or 'server' in error_str:
                    delay = base_delay * (attempt + 1)
                    print(f"  ⚠️ Server error. Retrying in {delay} seconds...")
                    time.sleep(delay)
                    continue
                if 'unauthorized' in error_str or 'invalid' in error_str:
                    print(f"  ❌ Authorization error when sending message: {e}")
                    return None

                # Non-recoverable error
                print(f"  ❌ Error sending message: {e}")
                if attempt == max_retries - 1:
                    import traceback
                    traceback.print_exc()
                return None

        print(f"  ❌ Failed to send after {max_retries} attempts")
        return None
    
    def process_prospect(self, row: pd.Series, create_draft_only: bool = True, test_mode: Optional[str] = None) -> bool:
        """
        Process a single prospect and send appropriate email
        
        Returns:
            bool: True if an email was sent, False otherwise
        """
        # Convert Series to scalar values safely
        company_name = str(row['company_name']) if not isinstance(row['company_name'], str) else row['company_name']
        first_name = str(row['first_name']) if not isinstance(row['first_name'], str) else row['first_name']
        email = str(row['email']) if not isinstance(row['email'], str) else row['email']
        
        # Handle NaN values
        if first_name.lower() in ['nan', 'none', '']:
            first_name = email.split('@')[0].capitalize()  # Use email username as fallback
        if company_name.lower() in ['nan', 'none', '']:
            company_name = email.split('@')[1].split('.')[0].capitalize()  # Use domain as fallback
        
        print(f"\n{'='*60}")
        print(f"Processing: {company_name} - {first_name} ({email})")
        print(f"{'='*60}")
        
        # --- DOMAIN-LEVEL SKIP CHECK ---
        # If we have already received a reply from this domain, skip all other
        # contacts from the same company without even checking Gmail.
        prospect_domain = self._get_domain(email)
        if prospect_domain and prospect_domain not in self.GENERIC_DOMAINS and prospect_domain in self.archived_domains:
            print(f"  🏢 Skipping {email} — domain '@{prospect_domain}' already replied.")
            # Archive this contact so they are also removed from CSV
            tracking_key = email
            if tracking_key in self.tracking_db:
                try:
                    self.archive_prospect(tracking_key, reason='domain_replied')
                except Exception:
                    pass
            return False
        
        # Get tracking data for this prospect
        tracking_key = email
        if tracking_key not in self.tracking_db:
            # NEW PROSPECT - Try to reconstruct from Gmail sent folder
            # (In case JSON was deleted but emails were already sent)
            if not test_mode:  # Don't reconstruct in test mode (always fresh)
                print(f"  🔍 Prospect not in tracking - checking Gmail sent folder...")
                
                # Try to reconstruct tracking data from Gmail
                # CRITICAL: Pass source_csv so we can remove from correct CSV file later
                source_csv = str(row.get('_source_csv', ''))
                reconstructed_data = self.reconstruct_tracking_from_gmail(email, company_name, first_name, source_csv)
                
                if reconstructed_data:
                    # Successfully reconstructed!
                    self.tracking_db[tracking_key] = reconstructed_data
                    self.save_tracking_db()
                    print(f"  💾 Tracking data RECONSTRUCTED and saved to {self.tracking_file}")
                    
                    # If they already replied, archive and stop
                    if reconstructed_data.get('received_reply'):
                        print(f"  ✓ Prospect already replied. Archiving and removing from CSV.")
                        try:
                            self.archive_prospect(tracking_key, reason="completed")
                        except Exception as e:
                            print(f"  ⚠️ Failed to archive: {e}")
                        return False
                    
                    # If all 3 emails sent, archive and stop
                    if len(reconstructed_data['emails_sent']) >= 3:
                        print(f"  ✓ All 3 emails already sent. Archiving and removing from CSV.")
                        try:
                            self.archive_prospect(tracking_key, reason="completed")
                        except Exception as e:
                            print(f"  ⚠️ Failed to archive: {e}")
                        return False
                    
                    # Continue processing with reconstructed data
                    print(f"  ➡️ Continuing from stage {len(reconstructed_data['emails_sent']) + 1}")
                else:
                    # No previous emails found - truly new prospect
                    print(f"  ✓ New prospect - no previous emails found")
                    self.tracking_db[tracking_key] = {
                        'company_name': company_name,
                        'first_name': first_name,
                        'email': email,
                        'emails_sent': [],
                        'received_reply': False,
                        'reply_at_stage': None,
                        '_source_csv': str(row.get('_source_csv', ''))
                    }
            else:
                # Test mode - always start fresh
                self.tracking_db[tracking_key] = {
                    'company_name': company_name,
                    'first_name': first_name,
                    'email': email,
                    'emails_sent': [],
                    'received_reply': False,
                    'reply_at_stage': None
                }
        
        prospect_data = self.tracking_db[tracking_key]
        
        # CHECK FOR BOUNCED EMAILS FIRST
        if prospect_data.get('emails_sent') and not test_mode:
            first_email = prospect_data['emails_sent'][0]
            thread_id = first_email.get('thread_id')
            
            # Check if email bounced
            if self.check_for_bounce(email, thread_id):
                print(f"  ❌ BOUNCE DETECTED - Email address invalid or unreachable")
                prospect_data['bounced'] = True
                prospect_data['bounce_detected_date'] = datetime.now().isoformat()
                self.save_tracking_db()
                
                # Archive as bounced
                try:
                    self.archive_prospect(tracking_key, reason="bounced")
                except Exception:
                    pass
                
                return False
        
        # TEST MODE: Send all 3 emails to test CSV prospects (for testing flow)
        if test_mode == "send_all_three":
            print(f"\n🧪 TEST MODE: Sending all 3 emails to {email}")
            
            # Track the first email's data for this test sequence
            first_email_msg_id = None
            first_thread_id = None
            
            for stage in [1, 2, 3]:
                print(f"\n📧 Sending Email {stage}...")
                
                # Get template
                template = EMAIL_TEMPLATES[stage]
                subject = template['subject'].format(company_name=company_name) if template['subject'] else None
                body = template['body'].format(company_name=company_name, first_name=first_name)
                
                # Create message
                if stage == 1:
                    message, first_email_msg_id = self.create_message_with_attachment(
                        to=email,
                        subject=subject,
                        body=body,
                        attachment_path=self.resume_path
                    )
                    message_id = self.send_message(message)
                    
                    if message_id:
                        # Get thread ID AND fetch the ACTUAL Message-ID that Gmail assigned
                        sent_msg = self.service.users().messages().get(
                            userId='me',
                            id=message_id,
                            format='metadata',
                            metadataHeaders=['Message-ID', 'Message-Id']
                        ).execute()
                        thread_id = sent_msg.get('threadId')
                        
                        # Extract the ACTUAL Message-ID header Gmail assigned (not our generated one)
                        headers = sent_msg.get('payload', {}).get('headers', [])
                        actual_message_id = None
                        for h in headers:
                            if h['name'] in ['Message-ID', 'Message-Id']:
                                actual_message_id = h['value']
                                break
                        
                        # Use Gmail's Message-ID, fallback to our generated one if not found
                        first_email_msg_id = actual_message_id if actual_message_id else first_email_msg_id
                        first_thread_id = thread_id  # The thread ID from Gmail
                        
                        print(f"  ✅ Email {stage} sent successfully!")
                        print(f"  📧 Message-ID: {first_email_msg_id}")
                        print(f"  📊 Thread ID: {first_thread_id}")
                        
                        # Save for next emails (use Gmail's actual Message-ID)
                        prospect_data['emails_sent'].append({
                            'stage': stage,
                            'message_id': message_id,
                            'thread_id': thread_id,
                            'email_message_id': first_email_msg_id,  # Store Gmail's actual Message-ID
                            'sent_date': datetime.now().isoformat()
                        })
                        
                        # Wait longer to let Gmail process the email and establish threading
                        if stage < 3:
                            print(f"  ⏳ Waiting 5 seconds for Gmail to process...")
                            time.sleep(5)
                    else:
                        print(f"  ❌ Failed to send Email {stage}")
                        return
                else:
                    # Follow-up in same thread
                    # Use the Message-ID and thread_id from Email 1 of THIS sequence
                    
                    # Get the original subject from Email 1
                    original_subject = EMAIL_TEMPLATES[1]['subject'].format(company_name=company_name)
                    reply_subject = f"Re: {original_subject}"
                    
                    message = self.create_reply_message(
                        to=email,
                        body=body,
                        thread_id=first_thread_id,  # Use the thread from Email 1 of THIS test
                        message_id=first_email_msg_id,  # Use the Message-ID from Email 1 of THIS test
                        subject=reply_subject
                    )
                    message_id = self.send_message(message)
                    
                    if message_id:
                        sent_msg = self.service.users().messages().get(
                            userId='me',
                            id=message_id,
                            format='minimal'
                        ).execute()
                        thread_id = sent_msg.get('threadId')
                        
                        # For follow-ups, store the Message-ID from Email 1 of THIS sequence
                        
                        prospect_data['emails_sent'].append({
                            'stage': stage,
                            'message_id': message_id,
                            'thread_id': thread_id,
                            'email_message_id': first_email_msg_id,  # Store Email 1's Message-ID
                            'sent_date': datetime.now().isoformat()
                        })
                        print(f"  ✅ Email {stage} sent successfully!")
                        print(f"  📊 Thread ID: {thread_id}")
                        
                        # Wait longer to let Gmail process
                        if stage < 3:
                            print(f"  ⏳ Waiting 5 seconds for Gmail to process...")
                            time.sleep(5)
                    else:
                        print(f"  ❌ Failed to send Email {stage}")
                        return False
            
            self.save_tracking_db()
            print("\n✅ All 3 test emails sent successfully!")
            return True
        
        # DRAFT MODE: Skip all checks, always create draft for current stage
        if create_draft_only:
            # Determine stage based on what's already sent
            current_stage = len(prospect_data['emails_sent']) + 1
            
            if current_stage > 3:
                print(f"  ℹ️ Note: Already created drafts for all 3 stages. Creating Stage {current_stage % 3 or 3} again...")
                current_stage = (current_stage - 1) % 3 + 1  # Cycle through 1, 2, 3
            
            print(f"  📧 DRAFT MODE - Creating Email {current_stage} draft...")
        else:
            # SEND MODE: Follow normal logic with checks
            # Determine which email stage we're at
            current_stage = len(prospect_data['emails_sent']) + 1
            
            # ALWAYS check for replies, even on first email (in case of reruns)
            if current_stage == 1:
                print(f"  📧 First contact - Preparing Email 1...")
                # Quick check: Do they already have emails from previous runs?
                if self.check_for_reply(email, None):
                    print(f"  ✅ REPLY ALREADY EXISTS from {company_name}!")
                    print(f"  💡 They replied before (possibly from previous run)")
                    prospect_data['received_reply'] = True
                    prospect_data['reply_at_stage'] = 0  # Before we could send
                    prospect_data['reply_detected_date'] = datetime.now().isoformat()
                    self.save_tracking_db()
                    # Archive prospect AND all colleagues from same domain
                    try:
                        self.archive_prospect(tracking_key, reason="completed")
                        self.archive_domain(email, reason='replied')
                    except Exception:
                        pass
                    return False
            else:
                # Only check for replies if we've already sent emails
                if prospect_data['received_reply']:
                    reply_stage = prospect_data.get('reply_at_stage')
                    if reply_stage:
                        print(f"  ✓ Already received reply after Email {reply_stage} from {company_name}. Skipping.")
                    else:
                        # Stage not recorded (from old tracking data)
                        last_sent = len(prospect_data['emails_sent'])
                        print(f"  ✓ Already received reply from {company_name} (after Email {last_sent}). Skipping.")
                    
                    # Archive prospect to remove from CSV and tracking
                    try:
                        self.archive_prospect(tracking_key, reason="completed")
                    except Exception as e:
                        print(f"  ⚠️ Failed to archive replied prospect: {e}")
                    
                    return False
                
                # Check for new reply (both in thread and general)
                # Get thread_id safely with null checks
                thread_id = None
                if prospect_data['emails_sent']:
                    last_email = prospect_data['emails_sent'][-1]
                    thread_id = last_email.get('thread_id')
                
                print(f"  🔍 Checking for replies (Thread ID: {thread_id or 'None'})...")
                
                # ALSO check for bounces in the thread
                if self.check_for_bounce(email, thread_id):
                    print(f"  ❌ BOUNCE DETECTED - Email address invalid or unreachable")
                    prospect_data['bounced'] = True
                    prospect_data['bounce_detected_date'] = datetime.now().isoformat()
                    self.save_tracking_db()
                    
                    # Archive as bounced
                    try:
                        self.archive_prospect(tracking_key, reason="bounced")
                    except Exception:
                        pass
                    return False
                
                if self.check_for_reply(email, thread_id):
                    # Calculate at which stage they replied
                    last_sent_stage = prospect_data['emails_sent'][-1]['stage'] if prospect_data['emails_sent'] else 1
                    
                    print(f"  ✅ REPLY DETECTED from {company_name}!")
                    print(f"  📊 Reply received after Email Stage {last_sent_stage}")
                    
                    prospect_data['received_reply'] = True
                    prospect_data['reply_at_stage'] = last_sent_stage
                    prospect_data['reply_detected_date'] = datetime.now().isoformat()
                    
                    self.save_tracking_db()
                    print(f"  💾 Tracking updated: No more emails will be sent.")
                    # Archive prospect AND all colleagues from same domain
                    try:
                        self.archive_prospect(tracking_key, reason="completed")
                        self.archive_domain(email, reason='replied')
                    except Exception:
                        pass
                    return False
            
            # Check if all 3 emails already sent
            if current_stage > 3:
                print(f"  ✓ All 3 emails already sent to {company_name}. No further action.")
                print(f"  📊 Emails sent: {len(prospect_data['emails_sent'])}, Current stage would be: {current_stage}")
                # Archive completed prospect
                try:
                    self.archive_prospect(tracking_key, reason="completed")
                except Exception:
                    pass
                return False
            
            # Extra safety check: Verify we haven't already sent this stage
            already_sent_stages = [e['stage'] for e in prospect_data['emails_sent']]
            if current_stage in already_sent_stages:
                print(f"  ⚠️ Email {current_stage} already sent to {company_name}! Skipping to prevent duplicate.")
                print(f"  📊 Already sent stages: {already_sent_stages}")
                return False
            
            # Check timing for follow-ups
            if current_stage > 1:
                last_email = prospect_data['emails_sent'][-1]
                last_sent_date = datetime.fromisoformat(last_email['sent_date'])
                days_since = (datetime.now() - last_sent_date).days
                required_days = TIMING_CONFIG[current_stage] - TIMING_CONFIG[current_stage - 1]
                
                if days_since < required_days:
                    print(f"  ⏳ Too soon for Email {current_stage}. Need {required_days} days, only {days_since} days passed.")
                    return False
                
                print(f"  📧 Follow-up time - Preparing Email {current_stage}...")
        
        # Get template
        template = EMAIL_TEMPLATES[current_stage]
        subject = template['subject'].format(company_name=company_name) if template['subject'] else None
        body = template['body'].format(company_name=company_name, first_name=first_name)
        
        # Safety check: ensure body is not empty
        if not body or not body.strip():
            print(f"  ❌ ERROR: Email body is empty! Template: {template['body'][:50]}...")
            print(f"  ❌ company_name={company_name}, first_name={first_name}")
            return False
        
        print(f"  📝 Email body preview: {body[:100]}...")
        
        # Create message
        if current_stage == 1:
            # First email with attachment
            message, generated_msg_id = self.create_message_with_attachment(
                to=email,
                subject=subject,
                body=body,
                attachment_path=self.resume_path
            )
        else:
            # Follow-up in same thread
            # ALWAYS use the Message-ID from Email 1 for proper threading
            if not prospect_data['emails_sent']:
                print(f"  ❌ Error: No previous emails found for follow-up. Skipping.")
                return False
            
            first_email_data = prospect_data['emails_sent'][0]
            
            # Get the original subject from Email 1
            original_subject = EMAIL_TEMPLATES[1]['subject'].format(company_name=company_name)
            reply_subject = f"Re: {original_subject}"
            
            # Use the Message-ID from Email 1 (the first email in the thread)
            email_message_id = first_email_data.get('email_message_id', first_email_data.get('message_id'))
            thread_id = first_email_data.get('thread_id')
            
            # VALIDATION: Check if we have valid threading data
            if not thread_id or not email_message_id:
                print(f"  ⚠️ Warning: Missing thread_id or message_id for follow-up")
                print(f"     thread_id: {thread_id}, message_id: {email_message_id}")
                print(f"  ⚠️ FALLBACK: Sending as new email instead of reply to avoid errors")
                
                # Send as new email (not a reply)
                message, generated_msg_id = self.create_message_with_attachment(
                    to=email,
                    subject=subject,
                    body=body,
                    attachment_path=None  # No attachment for follow-ups
                )
            else:
                # Normal reply with threading
                try:
                    message = self.create_reply_message(
                        to=email,
                        body=body,
                        thread_id=thread_id,
                        message_id=email_message_id,  # Always use Email 1's Message-ID
                        subject=reply_subject
                    )
                except Exception as threading_error:
                    print(f"  ⚠️ Error creating threaded reply: {threading_error}")
                    print(f"  ⚠️ FALLBACK: Sending as new email instead")
                    
                    # Fallback to new email
                    message, generated_msg_id = self.create_message_with_attachment(
                        to=email,
                        subject=subject,
                        body=body,
                        attachment_path=None
                    )
        
        # Create draft or send
        if create_draft_only:
            draft_id = self.create_draft(message)
            if draft_id:
                print(f"  ✅ Draft created for Email {current_stage}!")
                # Note: We don't update tracking for drafts yet
            return False  # Drafts are not sent, so no jitter needed
        else:
            # SEND MODE
            print(f"  🚀 SEND MODE ACTIVE - Sending Email {current_stage}...")

            # Safety: prevent duplicate sends if last email was sent moments ago
            if prospect_data.get('emails_sent'):
                try:
                    last_sent = prospect_data['emails_sent'][-1]
                    last_sent_date = datetime.fromisoformat(last_sent['sent_date'])
                    seconds_since = (datetime.now() - last_sent_date).total_seconds()
                    if seconds_since < 60:
                        print(f"  ⚠️ Last email was sent {int(seconds_since)}s ago. Skipping to avoid duplicate sends.")
                        return False
                except Exception:
                    pass

            # If this is a follow-up, double-check thread hasn't received a reply already
            followup_thread_id = None
            if current_stage > 1 and prospect_data.get('emails_sent'):
                followup_thread_id = prospect_data['emails_sent'][0].get('thread_id')
                if followup_thread_id and self.check_for_reply(email, followup_thread_id):
                    print(f"  ✓ Reply already present in thread {followup_thread_id}. Skipping send.")
                    prospect_data['received_reply'] = True
                    prospect_data['reply_at_stage'] = prospect_data['emails_sent'][-1]['stage'] if prospect_data['emails_sent'] else current_stage-1
                    prospect_data['reply_detected_date'] = datetime.now().isoformat()
                    self.save_tracking_db()
                    # Archive prospect AND all colleagues from same domain
                    try:
                        self.archive_prospect(tracking_key, reason="completed")
                        self.archive_domain(email, reason='replied')
                    except Exception:
                        pass
                    return False

            # Send the message
            message_id = self.send_message(message)
            if message_id:
                print(f"  📨 Getting thread information...")
                # Get thread ID AND the ACTUAL Message-ID that Gmail assigned
                sent_msg = self.service.users().messages().get(
                    userId='me',
                    id=message_id,
                    format='metadata',
                    metadataHeaders=['Message-ID', 'Message-Id']
                ).execute()

                thread_id = sent_msg.get('threadId')

                # Extract the ACTUAL Message-ID header Gmail assigned
                headers = sent_msg.get('payload', {}).get('headers', [])
                actual_message_id = None
                for h in headers:
                    if h['name'] in ['Message-ID', 'Message-Id']:
                        actual_message_id = h['value']
                        break

                # Normalize stored Message-ID (ensure angle brackets)
                def norm_mid(mid):
                    if not mid:
                        return mid
                    mid = str(mid).strip()
                    if mid.startswith('<') and mid.endswith('>'):
                        return mid
                    if '@' in mid:
                        return f"<{mid}>"
                    return mid

                # For stage 1, use Gmail's actual Message-ID
                if current_stage == 1:
                    email_message_id = norm_mid(actual_message_id) if actual_message_id else None
                    # If not found use generated_msg_id if available
                    if not email_message_id and 'generated_msg_id' in locals():
                        email_message_id = norm_mid(generated_msg_id)
                else:
                    # For follow-ups, use the Message-ID from first email
                    first_email_data = prospect_data['emails_sent'][0]
                    email_message_id = first_email_data.get('email_message_id') or first_email_data.get('message_id') or norm_mid(actual_message_id)
                    email_message_id = norm_mid(email_message_id)

                print(f"  💾 Saving to tracking database...")
                print(f"  📧 Message-ID: {email_message_id}")

                prospect_data['emails_sent'].append({
                    'stage': current_stage,
                    'message_id': message_id,
                    'thread_id': thread_id,
                    'email_message_id': email_message_id,  # Store normalized Message-ID
                    'sent_date': datetime.now().isoformat()
                })

                self.save_tracking_db()
                print(f"  ✅ Email {current_stage} sent successfully!")
                print(f"  📊 Tracking: Stage {current_stage}, Thread: {thread_id}")
                # If this was the final email (stage 3), archive prospect so it won't be checked again
                try:
                    if current_stage >= 3 and not test_mode:
                        self.archive_prospect(tracking_key)
                except Exception:
                    pass
                return True
            else:
                print(f"  ❌ Failed to send Email {current_stage}")
                return False
    
    def run(self, create_draft_only: bool = True, test_mode: Optional[str] = None):
        """
        Run the automation for all prospects with rate limiting
        
        Args:
            create_draft_only: If True, create drafts. If False, send emails.
            test_mode: "send_all_three" to send all 3 emails in sequence (testing)
        """
        if test_mode:
            print("\n" + "="*60)
            print(f"🧪 TEST MODE - Send All 3 Emails Sequentially")
            print("="*60)
            print(f"📊 Testing with {len(self.prospects)} prospects")
            print("="*60)
        else:
            print("\n" + "="*60)
            print(f"COLD EMAIL AUTOMATION - {'DRAFT MODE' if create_draft_only else 'SEND MODE'}")
            print("="*60)
            print(f"📊 Processing {len(self.prospects)} prospects")
            print(f"⏱️ Rate limit: Pause after every {RATE_LIMIT_EMAILS} emails")
            print("="*60)
        
        processed_count = 0
        
        for idx, row in self.prospects.iterrows():
            try:
                email_was_sent = self.process_prospect(row, create_draft_only, test_mode)
                processed_count += 1
                
                # Rate limiting: Wait after every 50 emails (not in test mode)
                if not test_mode and processed_count > 0 and processed_count % RATE_LIMIT_EMAILS == 0:
                    print(f"\n⏸️ RATE LIMIT: Processed {processed_count} emails")
                    print(f"⏳ Waiting {RATE_LIMIT_WAIT} seconds to respect Gemini API limits...")
                    time.sleep(RATE_LIMIT_WAIT)
                    print("✅ Resuming...\n")

                # Jitter between sends ONLY if email was actually sent (not just checking)
                if email_was_sent and not test_mode and not create_draft_only:
                    jitter = random.uniform(2.0, 5.0)
                    print(f"  ⏱️ Waiting {jitter:.1f}s before next send to avoid sending bursts...")
                    time.sleep(jitter)
                
            except Exception as e:
                print(f"  ❌ Error processing {row['company_name']}: {e}")
        
        print("\n" + "="*60)
        print(f"AUTOMATION COMPLETE - Processed {processed_count} prospects")
        print("="*60)


def main():
    """Main entry point"""
    print("="*70)
    print("COLD EMAIL AUTOMATION - STARTUP")
    print("="*70)
    
    # Step 1: Check Gmail account preference
    TOKEN_FILE = "token.json"
    
    if os.path.exists(TOKEN_FILE):
        # Try to read existing account info
        try:
            # Load credentials to show current email
            from google.oauth2.credentials import Credentials
            creds = Credentials.from_authorized_user_file(TOKEN_FILE)
            
            # Get email from token (if available)
            print("\n� Gmail Account Status")
            print("   ✓ Currently authenticated")
            
            # Try to get email info
            try:
                from googleapiclient.discovery import build
                service = build('gmail', 'v1', credentials=creds)
                profile = service.users().getProfile(userId='me').execute()
                current_email = profile.get('emailAddress', 'Unknown')
                print(f"   ✓ Email: {current_email}")
            except:
                print("   ✓ Authentication token found")
            
        except:
            print("\n📧 Gmail Account Status")
            print("   ⚠️ Token file exists but may be expired")
        
        print("\n" + "-"*70)
        print("GMAIL ACCOUNT OPTIONS:")
        print("-"*70)
        print("1. Continue with same Gmail account (use existing login)")
        print("2. Switch to different Gmail account (login with new account)")
        print("-"*70)
        
        choice = input("\nEnter your choice (1 or 2): ").strip()
        
        if choice == "2":
            print("\n🔄 Switching Gmail account...")
            os.remove(TOKEN_FILE)
            print("✅ Old authentication deleted")
            print("\n🔐 You will now login with a different Gmail account")
            print("   → A browser window will open")
            print("   → Sign in with your desired Gmail account")
            print("   → Allow access to Gmail API")
            input("\nPress ENTER to continue and authenticate...")
        else:
            print("\n✅ Using existing Gmail account")
    
    else:
        # No token exists - first time setup
        print("\n🔐 Gmail Authentication Required (First Time)")
        print("   → A browser window will open")
        print("   → Sign in with your Gmail account")
        print("   → Allow access to Gmail API")
        print("="*70)
        
        input("\nPress ENTER to continue and authenticate...")
    
    # Configuration - support a mail directory with multiple CSVs (mail/1.csv, 2.csv, ...)
    MAIL_DIR = "mail"
    EXCEL_FILE = None
    RESUME_PATH = "Gaurav_Resume.pdf"

    if os.path.isdir(MAIL_DIR):
        # Use the directory as the source; load_prospects will concatenate CSVs inside
        csv_files = sorted([f for f in os.listdir(MAIL_DIR) if f.lower().endswith('.csv')])
        if not csv_files:
            print(f"\n❌ Error: No CSV files found in '{MAIL_DIR}' directory!")
            print("Please add one or more CSV files like mail/1.csv, mail/2.csv")
            return

        print(f"\n✓ Found {len(csv_files)} CSV file(s) in '{MAIL_DIR}': {csv_files}")

        # Auto-clean each CSV file in the directory
        for fname in csv_files:
            path = os.path.join(MAIL_DIR, fname)
            print(f"\n→ Preparing: {path}")
            if not auto_clean_csv(path):
                print(f"❌ Could not prepare CSV: {path}. Please fix and retry.")
                return

        EXCEL_FILE = MAIL_DIR
    else:
        # Fallback to single CSV file (legacy)
        EXCEL_FILE = "mail.csv"
        if not os.path.exists(EXCEL_FILE):
            print(f"\n❌ Error: {EXCEL_FILE} not found!")
            print("Please create mail.csv with columns: Name, Email, Company (any format)")
            return

        print(f"\n✓ Found: {EXCEL_FILE}")

        # AUTO-CLEAN CSV (fix headings & extract first names)
        if not auto_clean_csv(EXCEL_FILE):
            print("❌ Could not prepare CSV. Please check the file format.")
            return
    
    if not os.path.exists(RESUME_PATH):
        print(f"⚠️ Warning: {RESUME_PATH} not found!")
        print("First emails will be sent without resume attachment.")
    
    # Initialize automation (this will trigger Gmail auth if needed)
    # Auto-recovery will handle any JSON issues automatically
    try:
        automation = ColdEmailAutomation(
            excel_file=EXCEL_FILE,
            resume_path=RESUME_PATH
        )
    except Exception as e:
        print(f"\n❌ Error initializing automation: {e}")
        print("Please check your configuration and try again.")
        return
    
    # USER CHOICE: Draft or Send
    print("\n" + "="*60)
    print("📧 EMAIL CAMPAIGN OPTIONS")
    print("="*60)
    print("1. Create drafts only (safe mode - review before sending)")
    print("2. Send emails directly (live mode - sends immediately)")
    print("="*60)
    
    choice = input("\nEnter your choice (1 or 2): ").strip()
    
    # Draft mode - ask for sub-option
    if choice == "1":
        print("\n✅ DRAFT MODE SELECTED")
        print("\n" + "-"*60)
        print("DRAFT MODE OPTIONS:")
        print("-"*60)
        print("1. Create single draft for next email stage")
        print("2. TEST: Send all 3 emails (1, 2, 3) sequentially to test.csv")
        print("-"*60)
        
        draft_choice = input("\nEnter your choice (1 or 2): ").strip()
        
        if draft_choice == "2":
            # Test mode with test.csv
            TEST_CSV = "test.csv"
            if not os.path.exists(TEST_CSV):
                print(f"\n❌ Error: {TEST_CSV} not found!")
                print("Please create test.csv with same format as mail.csv")
                print("This file should contain email addresses for testing only.")
                return
            
            print(f"\n✓ Found: {TEST_CSV}")
            
            # Clean test CSV
            if not auto_clean_csv(TEST_CSV):
                print("❌ Could not prepare test CSV.")
                return
            
            # Load test CSV instead with TEST MODE flag
            automation_test = ColdEmailAutomation(
                excel_file=TEST_CSV,
                resume_path=RESUME_PATH,
                is_test_mode=True  # Use email_tracking_test.json
            )
            
            print("\n⚠️ WARNING: This will SEND all 3 emails (1, 2, 3) to test.csv addresses!")
            print(f"📁 Tracking file: email_tracking_test.json (separate from production)")
            confirm = input("Are you sure? (yes/no): ").strip().lower()
            if confirm != "yes":
                print("❌ Cancelled.")
                return
            
            print("✅ TEST MODE: Sending all 3 emails sequentially...")
            automation_test.run(create_draft_only=False, test_mode="send_all_three")
            return
        else:
            print("✅ DRAFT MODE: Creating single draft for next stage.")
            automation.run(create_draft_only=True)
            return
    
    # Send mode
    create_draft_only = False
    confirm = input("\n⚠️ WARNING: This will SEND emails directly! Are you sure? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("❌ Cancelled. Switching to draft mode for safety.")
        create_draft_only = True
    else:
        print("✅ LIVE MODE: Emails will be sent directly!")
    
    # Run automation
    automation.run(create_draft_only=create_draft_only)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrupted by user. Exiting gracefully.")
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
