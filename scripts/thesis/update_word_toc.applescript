on run argv
    if (count of argv) is 0 then error "Thiếu đường dẫn DOCX"
    set docPath to item 1 of argv
    set macPath to (POSIX file docPath) as text
    with timeout of 600 seconds
        tell application "Microsoft Word"
            activate
            open file name macPath read only false add to recent files false
            delay 3
            set thesisDoc to active document
            set currentTOC to table of contents 1 of thesisDoc
            update currentTOC
            update page numbers currentTOC
            repaginate thesisDoc
            set saved of thesisDoc to false
            save thesisDoc
            close thesisDoc saving yes
        end tell
    end timeout
end run
