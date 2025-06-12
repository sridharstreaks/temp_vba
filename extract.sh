awk '
  /<util:set id="some_word"/ { in=1 }
  in && /<value>/ {
    gsub(/.*<value>|<\/value>.*/, "", $0)
    printf "\"%s\",", $0
  }
  /<\/util:set>/ { in=0 }
' file.xml | sed 's/,$//'
