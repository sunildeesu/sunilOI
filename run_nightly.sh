#!/bin/bash
#
# Nightly driver: regenerate the participant-OI Excel report, then commit and
# push the refreshed workbook so the copy on GitHub stays current.
#
# Invoked by the LaunchAgent com.marketanalysis.participantoi at 22:30 IST.
# Absolute tool paths are used because launchd runs with a minimal PATH.

set -u

REPO="/Users/sunildeesu/MarketAnalysisOI"
PY="/usr/bin/python3"
GIT="/usr/bin/git"

cd "$REPO" || { echo "cannot cd to $REPO"; exit 1; }

echo "===== nightly run $(date '+%Y-%m-%d %H:%M:%S %Z') ====="

# 1. Build the report ------------------------------------------------------- #
"$PY" "$REPO/participant_oi.py"
status=$?
if [ "$status" -ne 0 ]; then
    echo "participant_oi.py failed (exit $status); skipping git commit/push"
    exit "$status"
fi

# 2. Commit & push the workbook if it changed ------------------------------- #
if [ -n "$("$GIT" status --porcelain -- participant_oi.xlsx)" ]; then
    "$GIT" add participant_oi.xlsx
    "$GIT" commit -m "Update participant OI report - $(date '+%Y-%m-%d')"
    if "$GIT" push origin main; then
        echo "pushed updated report to GitHub"
    else
        echo "git push failed (keychain locked or offline?); will retry next run"
        exit 1
    fi
else
    echo "report unchanged; nothing to commit"
fi

echo "===== done ====="
