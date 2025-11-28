
(*
 *********************************************
 * /applescript/copyPresentation.applescript *
 *********************************************
*)

(*
  Copies slides from one Powerpoint presentation to another.
  
  If you want it to close the source file code "yes" as the
  final parameter.
*)

on copySlides(fromFile, toFile, afterSlide, closeFromFile)
	tell application "Microsoft PowerPoint"
		activate
		
		open fromFile
		open toFile
		
		set sourceSlides to slides of presentation fromFile
		set targetSlides to slides of presentation toFile
		
		-- Check whether the specified to insert after is valid
		set targetSlidesCount to length of targetSlides
		if afterSlide > targetSlidesCount then
			display dialog "Can't paste after slide " & afterSlide & ". Maximum slide number is " & targetSlidesCount & "."
			
			-- Can't continue
			return
		end if
		
		-- Copy over from source to destination - 1 slide at a time
		repeat with i from 1 to length of sourceSlides
			-- Copy a source slide to the clipboard
			set thisSlide to item i of sourceSlides
			copy object thisSlide
			
			-- Select the slide to paste after
			if afterSlide is 0 then
				select last slide of presentation toFile
			else
				select slide afterSlide of presentation toFile
				
				-- afterSlide incremeneted to be the slide just inserted
				set afterSlide to afterSlide + 1
				
			end if
			
			-- Paste slide from the clipboard
			tell active window
				set view type to slide sorter view
				paste object its view
				
				set view type to normal view
			end tell
			
			
		end repeat
		
		if closeFromFile is "yes" then
			my closeDocumentWindow(fromFile)
		end if
		
	end tell
end copySlides

on closeDocumentWindow(filename)
	
	tell application "Microsoft PowerPoint"
		set docWindows to every document window
		set foundWindow to "no"
		repeat with docWindow in docWindows
			set fn to full name of presentation of docWindow
			if fn is filename then
				set foundWindow to "yes"
				set foundWindowObject to docWindow
			end if
			
		end repeat
		
		if foundWindow is "yes" then
			close foundWindowObject
		end if
		
	end tell
end closeDocumentWindow

