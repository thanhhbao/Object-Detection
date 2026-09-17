on run argv
    if (count of argv) is 0 then error "Thiếu đường dẫn DOCX"
    set docPath to item 1 of argv
    set macPath to (POSIX file docPath) as text
    with timeout of 300 seconds
        tell application "Microsoft Word"
            activate
            open file name macPath read only false add to recent files false
            delay 2
            set thesisDoc to active document
            set fieldCount to count of fields of thesisDoc
            repeat with fieldIndex from fieldCount to 1 by -1
                update field (field fieldIndex of thesisDoc)
            end repeat
            set tocCount to count of tables of contents of thesisDoc
            repeat with tocIndex from 1 to tocCount
                set currentTOC to table of contents tocIndex of thesisDoc
                update currentTOC
                update page numbers currentTOC
            end repeat
            repaginate thesisDoc
            repeat with tocIndex from 1 to tocCount
                set currentTOC to table of contents tocIndex of thesisDoc
                update page numbers currentTOC
            end repeat
            save thesisDoc
            close thesisDoc saving yes
        end tell
    end timeout
end run
