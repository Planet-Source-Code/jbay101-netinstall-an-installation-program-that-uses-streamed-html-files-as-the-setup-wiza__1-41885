function browseControl(x,y)
   window.status="install::" & x & "," & y & ",browse"
   window.status=""
end function

function progressControl(x,y)
   window.status="install::" & x & "," & y & ",progress"
   window.status=""
end function

function startDownload(index,snext)
   window.status="install::" & index & "," & snext & ",copy"
   window.status=""
end function