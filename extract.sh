grep -oP '(?<=<util:set id="some_word"[^>]*>).*?(?=</util:set>)' file.xml \
  | grep -oP '(?<=<value>)[^<]+' \
  | sed 's/.*/"&"/' \
  | paste -sd, -
