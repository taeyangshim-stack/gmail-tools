# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script project for cleaning up duplicate/unnecessary emails in Gmail (taeyang.shim@gmail.com). The script runs in Google Apps Script (https://script.google.com), not locally — there is no build system, package manager, or test framework.

## Key File

- `gmail-cleanup.gs` — Single Apps Script file with all cleanup logic. Uses `GmailApp` API for search and batch trash operations.

## How to Deploy/Run

1. Copy `gmail-cleanup.gs` contents into a Google Apps Script project at https://script.google.com
2. Run functions individually: `step1_exosphere()` → `step2_outlook()` → `step3_duplicates()`, or `runAll()` for all at once
3. Gmail API has daily quota limits — if exceeded, wait 24 hours and re-run

## Architecture

Three-step pipeline, each targeting a different category of unwanted mail:

1. **step1_exosphere**: Bulk-deletes security alert spam from `no-reply@exosp.com` using search query + batch `moveThreadsToTrash()`
2. **step2_outlook**: Removes Microsoft Outlook test messages
3. **step3_duplicates**: Detects IMAP sync duplicates by composite key (subject + sender + timestamp rounded to minute). Processes in batches of 50 threads, caps at 2000 messages per run to stay within Apps Script execution time limits. Only trashes single-message threads where the message is a duplicate.

## Important Constraints

- `DRY_RUN` flag at top of file controls preview vs. actual deletion mode
- `moveThreadsToTrash()` is used for batch operations (1 API call per batch) to minimize quota usage
- Trash operations are reversible for 30 days
- `step3_duplicates` needs multiple runs if mailbox has >2000 messages (5-minute execution limit)
- Language: Korean comments and log messages throughout
