"""
Automated tests for domain-wise archive logic.

Tests run WITHOUT Gmail authentication by injecting a stub Gmail service and
a pre-populated tracking_db. No real emails are sent.

Run:
    python test_domain_archive.py
"""

import os
import sys

# Force UTF-8 stdout so emoji in cold_email_automation.py don't crash on
# Windows terminals using cp1252 / other narrow encodings.
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import json
import shutil
import tempfile
import traceback
import pandas as pd
from datetime import datetime
from unittest.mock import MagicMock, patch


# ---------------------------------------------------------------------------
# Minimal stub so we can import the class without real Google credentials
# ---------------------------------------------------------------------------

# Stub out network-dependent imports before importing the module
mock_gmail_service = MagicMock()
mock_gmail_service.users().getProfile(userId='me').execute.return_value = {
    'emailAddress': 'test_user@gmail.com'
}

# Patch get_google_credentials and build_gmail_service
google_creds_patch = patch(
    'cold_email_automation.get_google_credentials',
    return_value=MagicMock()
)
gmail_service_patch = patch(
    'cold_email_automation.build_gmail_service',
    return_value=mock_gmail_service
)

google_creds_patch.start()
gmail_service_patch.start()

from cold_email_automation import ColdEmailAutomation  # noqa: E402

# ---------------------------------------------------------------------------
# Helper: build a minimal ColdEmailAutomation object using a temp directory
# ---------------------------------------------------------------------------

def make_automation(tmp_dir: str, prospects: list) -> ColdEmailAutomation:
    """Create a ColdEmailAutomation pointing at a temp directory."""
    csv_path = os.path.join(tmp_dir, 'mail', '1.csv')
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    pd.DataFrame(prospects).to_csv(csv_path, index=False)

    # Patch load_prospects so we control data
    auto = ColdEmailAutomation.__new__(ColdEmailAutomation)
    auto.service = mock_gmail_service
    auto.authenticated_email = 'sender@gmail.com'
    auto.tracking_folder = os.path.join(tmp_dir, 'tracking')
    os.makedirs(auto.tracking_folder, exist_ok=True)
    auto.excel_file = os.path.join(tmp_dir, 'mail')
    auto.resume_path = 'fake_resume.pdf'
    auto.is_test_mode = False

    email_safe = 'sender_at_gmail_com'
    auto.tracking_file = os.path.join(auto.tracking_folder, f'email_tracking_{email_safe}.json')
    auto.archive_file = os.path.join(auto.tracking_folder, f'email_archive_{email_safe}.json')
    auto.archived_domains_file = os.path.join(auto.tracking_folder, f'archived_domains_{email_safe}.json')

    auto.tracking_db = {}
    auto.prospects = pd.DataFrame(prospects)
    auto.archived_domains = auto._load_archived_domains()
    return auto


def seed_tracking(auto: ColdEmailAutomation, email: str, company: str, name: str,
                  stage: int = 1, source_csv: str = '1.csv'):
    """Add a prospect to tracking_db as if an email had been sent."""
    auto.tracking_db[email] = {
        'company_name': company,
        'first_name': name,
        'email': email,
        'emails_sent': [{'stage': stage, 'message_id': 'msg1',
                         'thread_id': 'thread1', 'email_message_id': '<mid@x>',
                         'sent_date': datetime.now().isoformat()}],
        'received_reply': False,
        'reply_at_stage': None,
        '_source_csv': source_csv,
    }


# ---------------------------------------------------------------------------
# Test runner
# ---------------------------------------------------------------------------

PASSED = []
FAILED = []

def run_test(name, fn):
    tmp = tempfile.mkdtemp(prefix='email_test_')
    try:
        fn(tmp)
        PASSED.append(name)
        print(f'  [PASS]  {name}')
    except AssertionError as e:
        FAILED.append(name)
        print(f'  [FAIL]  {name}')
        print(f'        AssertionError: {e}')
    except Exception:
        FAILED.append(name)
        print(f'  [FAIL]  {name}')
        traceback.print_exc()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

def test_get_domain_normal(tmp):
    auto = make_automation(tmp, [])
    assert auto._get_domain('alice@startup.com') == 'startup.com'

def test_get_domain_uppercase(tmp):
    auto = make_automation(tmp, [])
    assert auto._get_domain('BOB@Company.IO') == 'company.io'

def test_get_domain_invalid(tmp):
    auto = make_automation(tmp, [])
    assert auto._get_domain('not-an-email') == ''
    assert auto._get_domain('') == ''

def test_save_and_load_archived_domains(tmp):
    auto = make_automation(tmp, [])
    auto._save_archived_domains('startup.com', reason='replied')
    auto._save_archived_domains('acme.org', reason='replied')

    # Reload from disk
    domains = auto._load_archived_domains()
    assert 'startup.com' in domains
    assert 'acme.org' in domains

def test_save_archived_domains_no_duplicates(tmp):
    auto = make_automation(tmp, [])
    auto._save_archived_domains('startup.com', reason='replied')
    auto._save_archived_domains('startup.com', reason='replied')  # duplicate

    with open(auto.archived_domains_file, encoding='utf-8') as f:
        data = json.load(f)
    count = sum(1 for e in data if e.get('domain') == 'startup.com')
    assert count == 1, f'Expected 1 entry, got {count}'

def test_archive_domain_archives_colleagues(tmp):
    """When alice replies, bob (same domain) should be archived."""
    prospects = [
        {'company_name': 'Startup', 'first_name': 'Alice', 'email': 'alice@startup.com'},
        {'company_name': 'Startup', 'first_name': 'Bob',   'email': 'bob@startup.com'},
    ]
    auto = make_automation(tmp, prospects)
    seed_tracking(auto, 'alice@startup.com', 'Startup', 'Alice')
    seed_tracking(auto, 'bob@startup.com',   'Startup', 'Bob')

    # Simulate alice replying — archive alice manually first, then call archive_domain
    auto.tracking_db.pop('alice@startup.com', None)  # already archived
    auto.archive_domain('alice@startup.com', reason='replied')

    assert 'bob@startup.com' not in auto.tracking_db, \
        "Bob should have been removed from tracking_db"
    assert 'startup.com' in auto.archived_domains, \
        "startup.com should be in archived_domains set"
    assert os.path.exists(auto.archived_domains_file), \
        "archived_domains JSON file should exist"

def test_archive_domain_skips_generic(tmp):
    """gmail.com is a generic domain — colleagues should NOT be batch-archived."""
    prospects = [
        {'company_name': 'PersonalA', 'first_name': 'Alice', 'email': 'alice@gmail.com'},
        {'company_name': 'PersonalB', 'first_name': 'Bob',   'email': 'bob@gmail.com'},
    ]
    auto = make_automation(tmp, prospects)
    seed_tracking(auto, 'alice@gmail.com', 'PersonalA', 'Alice')
    seed_tracking(auto, 'bob@gmail.com',   'PersonalB', 'Bob')

    auto.archive_domain('alice@gmail.com', reason='replied')

    assert 'bob@gmail.com' in auto.tracking_db, \
        "Bob (different person, generic domain) should NOT be archived"
    assert 'gmail.com' not in auto.archived_domains

def test_archive_domain_does_not_archive_replier_twice(tmp):
    """archive_domain should only look at OTHER contacts, not the replier."""
    prospects = [
        {'company_name': 'Startup', 'first_name': 'Alice', 'email': 'alice@startup.com'},
    ]
    auto = make_automation(tmp, prospects)
    seed_tracking(auto, 'alice@startup.com', 'Startup', 'Alice')

    # alice is still in tracking_db — archive_domain must not crash
    count = auto.archive_domain('alice@startup.com', reason='replied')
    # count is colleagues EXCLUDING the replier
    assert count == 0

def test_archive_domain_idempotent(tmp):
    """Calling archive_domain twice for the same domain should not error."""
    prospects = [
        {'company_name': 'Startup', 'first_name': 'Alice', 'email': 'alice@startup.com'},
        {'company_name': 'Startup', 'first_name': 'Bob',   'email': 'bob@startup.com'},
    ]
    auto = make_automation(tmp, prospects)
    seed_tracking(auto, 'alice@startup.com', 'Startup', 'Alice')
    seed_tracking(auto, 'bob@startup.com',   'Startup', 'Bob')

    auto.archive_domain('alice@startup.com', reason='replied')
    auto.archive_domain('alice@startup.com', reason='replied')  # second call, should no-op

    with open(auto.archived_domains_file, encoding='utf-8') as f:
        data = json.load(f)
    count = sum(1 for e in data if e.get('domain') == 'startup.com')
    assert count == 1, f'Expected 1 JSON entry for startup.com, got {count}'

def test_process_prospect_skips_archived_domain(tmp):
    """process_prospect should return False immediately for archived domains."""
    prospects = [
        {'company_name': 'Startup', 'first_name': 'Bob', 'email': 'bob@startup.com'},
    ]
    auto = make_automation(tmp, prospects)
    seed_tracking(auto, 'bob@startup.com', 'Startup', 'Bob')

    # Mark startup.com as already archived
    auto.archived_domains.add('startup.com')

    row = auto.prospects.iloc[0]
    result = auto.process_prospect(row, create_draft_only=True)

    assert result is False, "Should return False for archived domain prospects"

def test_process_prospect_allows_generic_domain(tmp):
    """process_prospect should NOT skip contacts on generic domains even if domain is 'archived'."""
    prospects = [
        {'company_name': 'PersonalB', 'first_name': 'Bob', 'email': 'bob@gmail.com'},
    ]
    auto = make_automation(tmp, prospects)

    # Even if someone tries to put gmail.com in archived_domains, the check should
    # respect GENERIC_DOMAINS and not skip
    auto.archived_domains.add('gmail.com')  # shouldn't matter

    # We don't want it to actually call Gmail API — just check the skip logic
    # Since 'gmail.com' is in GENERIC_DOMAINS, the skip check is bypassed.
    # process_prospect will proceed to Gmail logic. We can verify by checking
    # it does NOT return False immediately (it will hit reconstruct_tracking which
    # calls Gmail).  We mock that path.
    with patch.object(auto, 'reconstruct_tracking_from_gmail', return_value=None), \
         patch.object(auto, 'create_draft', return_value='draft_id_123'):
        row = auto.prospects.iloc[0]
        # The function should not short-circuit on domain check (returns False only
        # because create_draft_only=True makes it return False at the end of processing)
        result = auto.process_prospect(row, create_draft_only=True)
        # result can be True or False depending on draft creation; the key is it
        # should NOT return False because of the domain skip (it should reach the draft stage)
        # We verify that reconstruct_tracking_from_gmail was actually called (meaning
        # the domain check did not trigger an early exit)
        auto.reconstruct_tracking_from_gmail.assert_called_once()

def test_archived_domains_persists_across_instances(tmp):
    """Archived domains written by one instance should be seen by a new instance."""
    prospects = [
        {'company_name': 'Startup', 'first_name': 'Alice', 'email': 'alice@startup.com'},
        {'company_name': 'Startup', 'first_name': 'Bob',   'email': 'bob@startup.com'},
    ]
    auto1 = make_automation(tmp, prospects)
    seed_tracking(auto1, 'alice@startup.com', 'Startup', 'Alice')
    seed_tracking(auto1, 'bob@startup.com',   'Startup', 'Bob')
    auto1.archive_domain('alice@startup.com', reason='replied')

    # Create a second instance reusing the same tracking folder
    auto2 = make_automation(tmp, prospects)
    assert 'startup.com' in auto2.archived_domains, \
        "Second instance should load 'startup.com' from persisted JSON"


# ---------------------------------------------------------------------------
# Run all tests
# ---------------------------------------------------------------------------

TESTS = [
    ('_get_domain: normal email',                         test_get_domain_normal),
    ('_get_domain: uppercase/mixed case',                 test_get_domain_uppercase),
    ('_get_domain: invalid input',                        test_get_domain_invalid),
    ('_save_archived_domains + _load_archived_domains',   test_save_and_load_archived_domains),
    ('_save_archived_domains: no duplicates',             test_save_archived_domains_no_duplicates),
    ('archive_domain: archives colleagues',               test_archive_domain_archives_colleagues),
    ('archive_domain: skips generic domains (gmail.com)', test_archive_domain_skips_generic),
    ('archive_domain: does not archive replier twice',    test_archive_domain_does_not_archive_replier_twice),
    ('archive_domain: idempotent (double call)',          test_archive_domain_idempotent),
    ('process_prospect: skips archived domain',          test_process_prospect_skips_archived_domain),
    ('process_prospect: allows generic domain through',  test_process_prospect_allows_generic_domain),
    ('archived_domains: persists across instances',       test_archived_domains_persists_across_instances),
]

if __name__ == '__main__':
    print('\n' + '='*60)
    print('DOMAIN-WISE ARCHIVE  --  Automated Test Suite')
    print('='*60 + '\n')

    for name, fn in TESTS:
        run_test(name, fn)

    print('\n' + '-'*60)
    total = len(PASSED) + len(FAILED)
    print(f'Results: {len(PASSED)}/{total} passed')
    if FAILED:
        print('\nFailed tests:')
        for t in FAILED:
            print(f'  * {t}')
        sys.exit(1)
    else:
        print('\nAll tests passed!')
        sys.exit(0)
