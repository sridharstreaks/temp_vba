word=$(
  awk '
    # When we hit the opening tag, turn our flag on
    /<util:set[[:space:]]+id="some_word"/ { in_block = 1; next }

    # When we hit the closing tag, turn it off
    /<\/util:set>/            { in_block = 0; next }

    # Only inside that block, grab <value> lines
    in_block && /<value>/ {
      # strip the tags, leaving just the inner text
      gsub(/.*<value>|<\/value>.*/, "")
      # accumulate into vals with quotes+comma
      vals = vals "\"" $0 "\","
    }

    END {
      # drop trailing comma and print
      sub(/,$/, "", vals)
      print vals
    }
  ' "$file"
)
