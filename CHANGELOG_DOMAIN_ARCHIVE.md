# Domain-wise Archive Feature — Contribution Notes

## What Changed and Why

When multiple contacts from the same company are in the CSV, the old logic would continue sending reminder emails to all of them even after one person had already replied. This wastes outreach quota and looks unprofessional.

This contribution adds a **domain-level archiving** feature: as soon as a reply is detected from any contact, **every other active contact sharing the same email domain is automatically archived** and removed from future runs.

---

## Files Changed

### `cold_email_automation.py`

This is the only file that was modified. Below are all changes in logical order.

---

### 1. `__init__` — Load archived domains on startup

```diff
+        # Archived domains file — persists domains where a reply was received
+        self.archived_domains_file = os.path.join(
+            self.tracking_folder, f"archived_domains_{email_safe}.json"
+        )
+
         self.tracking_db = self.load_tracking_db()
+        self.archived_domains = self._load_archived_domains()
```

Also prints the archived domains on startup so the user can see which company domains are already skipped.

---

### 2. `GENERIC_DOMAINS` — Class-level constant

```python
GENERIC_DOMAINS = {
    'gmail.com', 'yahoo.com', 'yahoo.in', 'hotmail.com', 'outlook.com',
    'live.com', 'icloud.com', 'me.com', 'aol.com', 'protonmail.com',
    'proton.me', 'zoho.com', 'mail.com', 'ymail.com', 'googlemail.com',
}
```

Prevents batch-archiving unrelated people who happen to share a generic email provider. Domain archiving is only applied when the domain clearly belongs to a single organisation.

---

### 3. `_get_domain(email)` — New helper

Extracts the domain part of any email address (`user@company.com` -> `company.com`). Returns `''` on invalid input.

---

### 4. `_load_archived_domains()` — New helper

Reads `tracking/archived_domains_<account>.json` and returns a Python `set` of domain strings. Called once during `__init__`.

---

### 5. `_save_archived_domains(domain, reason)` — New helper

Atomically appends a new `{"domain": ..., "archived_at": ..., "reason": ...}` entry to the JSON file, preventing duplicates.

---

### 6. `archive_domain(replied_email, reason)` — New public method

Core of the feature. When called with the replying contact's email:

1. Extracts the domain.
2. Skips if it's a generic/personal domain or already archived.
3. Collects all remaining active contacts from `tracking_db` that share the domain.
4. Calls `archive_prospect` on each of them.
5. Persists the domain to `archived_domains.json`.

Returns the count of additional contacts archived.

---

### 7. `process_prospect` — Domain-level skip check (early exit)

Added at the **very top** of `process_prospect`, before any Gmail API calls:

```python
prospect_domain = self._get_domain(email)
if prospect_domain and prospect_domain not in self.GENERIC_DOMAINS \
        and prospect_domain in self.archived_domains:
    # Skip + archive from CSV immediately
    return False
```

This ensures contacts from already-responded domains are skipped cheaply without touching the Gmail API.

---

### 8. Reply detection paths now trigger `archive_domain`

All three code paths where a reply is detected now call `archive_domain` right after `archive_prospect`:

- Stage-1 pre-check (reply found before first email is due)
- Follow-up check (reply found before sending stage 2 or 3)
- Pre-send double-check (reply found just before a follow-up is sent)

```python
self.archive_prospect(tracking_key, reason="completed")
self.archive_domain(email, reason='replied')   # <- new
```

---

## New Tracking Files Created

| File | Description |
|---|---|
| `tracking/archived_domains_<account>.json` | Persistent list of domains that have replied. Auto-created on first use. |

---

## How It Works End-to-End

```
1. Run the script normally.
2. Reply detected from: alice@startup.com
3. alice@startup.com is archived (existing behaviour).
4. archive_domain("alice@startup.com") is called.
5. Domain "startup.com" is NOT in GENERIC_DOMAINS -> proceed.
6. All other contacts with @startup.com in tracking_db are archived + removed from CSV.
7. "startup.com" is saved to archived_domains_<account>.json.
8. Next run: bob@startup.com is encountered.
9. Early-exit check fires -> skipped instantly, no Gmail API call needed.
```

---

## Backwards Compatibility

- No changes to existing JSON tracking file formats.
- The new `archived_domains_*.json` file is auto-created; deleting it simply resets domain memory.
- All existing archive/tracking behaviour is preserved.
