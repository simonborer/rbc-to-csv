#!/bin/bash

# === Semantics are wrecked, so copy & paste cc data into a .txt file ===
# Script assumes you've filtered for only debits, no credits

if [ -z "$1" ]; then
    echo "Usage: $0 inputfile.txt" >&2
    exit 1
fi

INPUT="$1"

awk -v OFS=',' -v platform_date="$(uname)" '
BEGIN {
    # === Fix: Emit proper header ===
    print "\"Date\",\"Description\",\"\",\"Debit\",\"Credit\",\"\""
}
function to_iso_date(date_str, line_num, iso_date, cmd) {
    if (platform_date == "Darwin") {
        cmd = "date -j -f \"%b %d, %Y\" \"" date_str "\" +%Y-%m-%d"
    } else {
        cmd = "date -d \"" date_str "\" +%Y-%m-%d"
    }
    if ((cmd | getline iso_date) <= 0) {
        print "ERROR: Failed to convert date on line " line_num ": \"" date_str "\"" > "/dev/stderr"
        close(cmd)
        return ""
    }
    close(cmd)
    return iso_date
}
NR == 1 && $1 ~ /Date/i { next }

NF { lines[record_line++] = $0 }

END {
    i = 0
    while (i < record_line) {
        date_raw = lines[i++]
        desc1 = lines[i++]

        if (i < record_line && lines[i] ~ /^\$/) {
            amount_raw = lines[i++]
            desc2 = ""
        } else {
            desc2 = lines[i++]
            amount_raw = lines[i++]
        }

        # === Fix: Remove trailing whitespace that annoys BSD date ===
        gsub(/[[:space:]]+$/, "", date_raw)

        iso_date = to_iso_date(date_raw, i - 3)
        if (iso_date == "") continue

        gsub(/^[[:space:]]+|[[:space:]]+$/, "", desc1)
        gsub(/^[[:space:]]+|[[:space:]]+$/, "", desc2)
        description = desc1
        if (desc2 != "") description = description " " desc2

        gsub(/[^0-9.]/, "", amount_raw)

        # === Fix: Output with empty CSV columns ===
        print "\"" iso_date "\"", "\"" description "\"", "\"\"", "\"" amount_raw "\"", "\"\"", "\"\""
    }
}
' "$INPUT"
