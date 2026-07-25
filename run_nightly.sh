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

# 2. Commit & push only when the underlying data changed -------------------- #
# report_data.hash is a fingerprint of the source data (not the timestamp), so a
# holiday/weekend or a same-day re-run produces no diff here and no commit.
if [ -n "$("$GIT" status --porcelain -- report_data.hash)" ]; then
    "$GIT" add participant_oi.xlsx report_data.hash
    "$GIT" commit -m "Update participant OI report - $(date '+%Y-%m-%d')"
    if "$GIT" push origin main; then
        echo "pushed updated report to GitHub"
    else
        echo "git push failed (keychain locked or offline?); will retry next run"
        exit 1
    fi
else
    echo "data unchanged; nothing to commit"
    # discard the regenerated-but-identical workbook so the tree stays clean
    "$GIT" checkout -- participant_oi.xlsx 2>/dev/null || true
fi

echo "===== done ====="
