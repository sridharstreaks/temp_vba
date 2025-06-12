awk -v id="some_word" '
  # once we see <util:set … id="some_word" …>  up to </util:set>...
  $0 ~ "<util:set[^>]*id=\"" id "\"[^>]*>" , $0 ~ "</util:set>" {
    # if this line has a <value>…</value>, capture it
    if (match($0, /<value>([^<]+)<\/value>/, m)) {
      # print "value", and a comma
      printf "\"%s\",", m[1]
    }
  }
  END {
    # strip trailing comma & print newline
    sub(/,$/, "")
    print ""
  }
' input.xml
