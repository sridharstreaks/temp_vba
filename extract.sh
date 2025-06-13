awk '
  /<util:set[^>]*id="format_words"/ { inblock=1 }
  inblock && /<value>/ {
    gsub(/.*<value>|<\/value>.*/, "", $0)
    printf "\"%s\",", $0
  }
  /<\/util:set>/ && inblock { inblock=0; print "" }
' input.xml
