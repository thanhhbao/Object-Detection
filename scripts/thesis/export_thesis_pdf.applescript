-- Cập nhật toàn bộ field và mục lục, lưu DOCX, rồi xuất PDF.
--
-- Thứ tự bắt buộc: cập nhật field -> cập nhật mục lục -> phân trang lại ->
-- cập nhật số trang lần nữa -> lưu DOCX -> xuất PDF. Nếu xuất PDF trước khi
-- phân trang lại thì số trang in trong mục lục sẽ lệch so với nội dung.
--
-- Dùng:
--   osascript scripts/export_thesis_pdf.applescript <đường dẫn .docx> [<đường dẫn .pdf>]
-- Không truyền tham số thứ hai thì PDF được đặt cạnh DOCX, cùng tên.

on run argv
	if (count of argv) is 0 then error "Thiếu đường dẫn DOCX"
	set docPath to item 1 of argv
	if (count of argv) > 1 then
		set pdfPath to item 2 of argv
	else
		set AppleScript's text item delimiters to "."
		set parts to text items of docPath
		set parts to items 1 thru -2 of parts
		set pdfPath to ((parts as text) & ".pdf")
		set AppleScript's text item delimiters to ""
	end if

	set macDoc to (POSIX file docPath) as text
	set macPdf to (POSIX file pdfPath) as text

	with timeout of 900 seconds
		tell application "Microsoft Word"
			activate
			open file name macDoc read only false add to recent files false
			delay 2
			set theDoc to active document

			-- 1. cập nhật mọi field, duyệt ngược để field lồng nhau không bị bỏ sót
			set fieldCount to count of fields of theDoc
			repeat with i from fieldCount to 1 by -1
				update field (field i of theDoc)
			end repeat

			-- 2. cập nhật cả ba mục lục
			set tocCount to count of tables of contents of theDoc
			repeat with i from 1 to tocCount
				set aTOC to table of contents i of theDoc
				update aTOC
				update page numbers aTOC
			end repeat

			-- 3. phân trang lại rồi cập nhật số trang lần nữa
			repaginate theDoc
			repeat with i from 1 to tocCount
				update page numbers (table of contents i of theDoc)
			end repeat

			-- 4. lưu bản DOCX đã cập nhật
			save theDoc

			-- 5. xuất PDF
			save as theDoc file name macPdf file format format PDF

			-- "save as" đã trỏ document sang tệp PDF, nên đóng mà không lưu thêm
			close theDoc saving no
		end tell
	end timeout
	return pdfPath
end run
