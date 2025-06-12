#!/usr/bin/env bash
# file: extract_values_simple.sh

XML_FILE=$1
ID=$2

awk -v id="$ID" '
  # when we hit the <util:set id"…" line, start collecting…
  $0 ~ "<util:set id\"" id "\"" { inblock=1; next }
  # …until the closing tag
  $0 ~ "</util:set>"      { inblock=0 }
  # inside block, look for <value>…</value>
  inblock && match($0, /<value>([^<]+)<\/value>/, m) {
    vals[++n] = "\"" m[1] "\""
  }
  END {
    if (n>0) {
      for (i=1;i<=n;i++) {
        printf "%s%s", vals[i], (i<n?",":"")
      }
      printf "\n"
    }
  }
' "$XML_FILE"
