awk '
  /<util:set[[:space:]]+id="some_word"/ { in_block=1; next }
  /<\/util:set>/          { in_block=0 }
  in_block && /<value>/  {
    # extract just the text between <value>…</value>
    if ( match($0, /<value>([^<]+)<\/value>/, m) ) {
      printf "\"%s\",", m[1]
    }
  }
  END {
    # strip trailing comma and add newline
    sub(/,$/, "", out = "")
    # but since we printed directly, re-print everything from stdout
    # instead, you can pipe through sed as shown below
  }
' file.xml | sed 's/,$//'
