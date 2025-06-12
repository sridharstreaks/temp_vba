grep -zoP '<util:set id="some_word"[^>]*>.*?</util:set>' file.xml \
  | grep -oP '<value>\K[^<]+' \
  | sed 's/.*/"&"/' \
  | paste -sd, -
