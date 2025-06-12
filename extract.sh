word=$(
  awk '
    BEGIN {
      in_block = 0
      vals     = ""
    }
    # 1) opening-tag: must be a <util:set> with id="some_word" anywhere in it
    /<util:set[^>]*id="some_word"[^>]*>/ {
      in_block = 1
      next
    }
    # 2) closing-tag: turn it off when you see </util:set>
    /<\/util:set[[:space:]]*>/ {
      in_block = 0
      next
    }
    # 3) only inside that block, grab <value>…</value>
    in_block {
      if ( match($0, /<value>([^<]+)<\/value>/, m) ) {
        vals = vals "\"" m[1] "\","
      }
    }
    END {
      sub(/,$/, "", vals)
      print vals
    }
  ' "$file"
)
