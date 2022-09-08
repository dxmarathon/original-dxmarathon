wget --mirror --convert-links --force-directories -e robots=off --timestamping --no-if-modified-since https://dxmarathon.com


find dxmarathon.com -type f -name "*.htm" | while read file; do iconv -f MS-ANSI -t UTF-8 "$file" >"$file.2"; mv "$file.2" "$file"; done;
