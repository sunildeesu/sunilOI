#!/bin/bash
#
# Nightly driver: regenerate the participant-OI Excel report, then commit and
# push the refreshed workbook so the copy on GitHub stays current.
#
# Invoked by the LaunchAgent com.marketanalysis.participantoi twice each weekday:
# 20:30 IST (early) and 22:30 IST (fallback for late NSE publishing). The 22:30 run
# is a no-op when the 20:30 run already captured the data - commit is gated on a
# change to report_data.hash (see step 3 below).
# Absolute tool paths are used because launchd runs with a minimal PATH.

set -u

REPO="/Users/sunildeesu/MarketAnalysisOI"
NIX="/nix/var/nix/profiles/default/bin/nix"
GIT="/usr/bin/git"

cd "$REPO" || { echo "cannot cd to $REPO"; exit 1; }

echo "===== nightly run $(date '+%Y-%m-%d %H:%M:%S %Z') ====="

# 0. Reproducible Python via nix -------------------------------------------- #
# Build (or reuse) the pinned python3 + openpyxl + requests env. The --out-link
# is a GC root, so nix-collect-garbage won't remove it and cached rebuilds are
# near-instant and offline. This replaces the fragile Apple system python.
if ! "$NIX" build "$REPO#pythonEnv" --out-link "$REPO/.nix-python"; then
    echo "nix build of python env failed; aborting run"
    exit 1
fi
PY="$REPO/.nix-python/bin/python3"

# 1. Build the report ------------------------------------------------------- #
"$PY" "$REPO/participant_oi.py"
status=$?
if [ "$status" -ne 0 ]; then
    echo "participant_oi.py failed (exit $status); skipping git commit/push"
    exit "$status"
fi

# 2. Reconcile with GitHub before committing --------------------------------- #
# Merged PRs land on GitHub while this live clone keeps making nightly commits.
# Fetch and fast-forward so the next push is not rejected with "(fetch first)".
if ! "$GIT" fetch origin main; then
    echo "git fetch origin main failed; aborting commit/push"
    exit 1
fi
if ! "$GIT" merge --ff-only origin/main; then
    echo "git merge --ff-only origin/main failed (local and remote both diverged); attempting rebase"
    if ! "$GIT" rebase origin/main; then
        echo "git rebase origin/main failed; manual reconciliation needed"
        exit 1
    fi
fi

# 3. Commit & push only when the underlying data changed -------------------- #
# report_data.hash is a fingerprint of the source data (not the timestamp), so a
# holiday/weekend or a same-day re-run produces no diff here and no commit.
if [ -n "$("$GIT" status --porcelain -- report_data.hash)" ]; then
    "$GIT" add participant_oi.xlsx report_data.hash
    "$GIT" commit -m "Update participant OI report - $(date '+%Y-%m-%d')"
    push_out=$("$GIT" push origin main 2>&1)
    push_status=$?
    if [ "$push_status" -eq 0 ]; then
        echo "pushed updated report to GitHub"
    else
        echo "git push failed (exit $push_status): $push_out"
        exit 1
    fi
else
    echo "data unchanged; nothing to commit"
    # discard the regenerated-but-identical workbook so the tree stays clean
    "$GIT" checkout -- participant_oi.xlsx 2>/dev/null || true
fi

echo "===== done ====="
