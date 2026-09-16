#!/usr/bin/env bash
set -u

BASE="https://schedule-gen-from-email.vercel.app/api/extract-event"

run_test() {
  local label="$1"
  local email="$2"
  local payload
  payload=$(python3 -c "import json,sys; print(json.dumps({'model':'gpt-4o-mini','email':sys.argv[1]}))" "$email")
  echo "==================================================================="
  echo "  $label"
  echo "==================================================================="
  echo "EMAIL:"
  echo "$email"
  echo ""
  echo "RESPONSE:"
  curl -sS -X POST "$BASE" \
    -H "Content-Type: application/json" \
    -d "$payload" \
    -w "\n[HTTP %{http_code} in %{time_total}s, %{size_download} bytes]\n"
  echo ""
  echo ""
}

run_test "T1 - Same-day with explicit end time" \
"Reminder: Team standup tomorrow from 2:00 PM to 3:00 PM in the conference room."

run_test "T2 - Date range with 'each day' times (expect one per day, each timed)" \
"Engineering Conference Oct 10-12, 2026 from 9:00 AM to 5:00 PM each day at the Moscone Center, San Francisco."

run_test "T3 - Date range with no times (expect one whole_day entry per calendar day)" \
"The annual company offsite runs from October 5 to October 8, 2026 in Lake Tahoe."

run_test "T4 - Multi-week date range with no times (expect one whole_day entry per calendar day)" \
"Tech Festival from October 10 to October 20, 2026 at the convention center."

run_test "T5 - Two independent multi-day events (expect one per event, each with end_time)" \
"Hackathon Day 1: July 10 2026 from 1:00 PM to 5:00 PM.
Hackathon Day 2: July 11 2026 from 1:00 PM to 3:00 PM.
Both days in Siebel Center."

run_test "T6 - Date range with 'each day' phrasing (expect one timed entry per day)" \
"Workshop on Oct 5, 6, and 7, 2026. Each day runs 9:00 AM to 5:00 PM in Room 100."

run_test "T7 - Range with 'daily' times (expect one timed entry per day)" \
"Client visit from Monday Oct 5 to Wednesday Oct 7, 2026 at their NYC office. Daily meetings 10am-4pm."

run_test "T8 - After Hours Chicago (your real email, expect whole_day=false, end_time=19:30)" \
"Reminder: Join us for a night of networking in Chicago!

After Hours in Chicago is an informal career reception in Chicago for CS & ECE students and companies to connect, network and socialize in a less formal setting than a typical career fair.

Thursday, October 1, 2026 from 5:30 PM - 7:30 PM
200 S. Wacker Drive, 4th Floor, Chicago, IL 60606

The registration form above will close on September 22, 2026 at 11:59PM Central Time."
