'SystemUtil.Run "chrome.exe","https://wordpress.com/?appromo"
Set exc1 = CreateObject("Excel.Application")                                
Set wb1 = exc1.Workbooks.Open("C:\Users\Balaji Vignesh\Desktop\framework.xlsx")  
Set sh1 = wb1.sheets(1)  
rc1=sh1.usedrange.rows.count
Set sh2 = wb1.sheets(2)  
rc2=sh2.usedrange.rows.count
For k= 2 To rc1 Step 1
	ctrl1=sh1.cells(k,1)
	If ctrl1=1 Then
		fid1=sh1.cells(k,2)
	
	For j = 2 To rc2 Step 1
		fid2=sh2.cells(j,1)
		If fid2=fid1 Then
			ctrl2=sh2.cells(j,2)
			If ctrl2=1 Then
				sfid=sh2.cells(j,3)
				If sfid="SF1" Then

	set ExcelObj = createobject("excel.application")
	ExcelObj.visible = True
	set WorkbookObj= ExcelObj.Workbooks.Open ("C:\Users\Balaji Vignesh\Desktop\Data_sheet1.xlsx")

	set SheetObj = WorkbookObj.Worksheets("second innings")

	Row=SheetObj.UsedRange.Rows.Count
		For  i= 1 to Row step 1

			LoginUserId=SheetObj.cells(i,1).value

			Password=SheetObj.cells(i,2).value

			Browser("WordPress.com").Page("WordPress.com").Link("Sign In").Click
			Browser("WordPress.com").Page("WordPress.com ‹ Log In").WebEdit("log").Set LoginUserId
			wait(1)
			Browser("WordPress.com").Page("WordPress.com ‹ Log In").WebEdit("pwd").Set Password
			wait(1)
			Browser("WordPress.com").Page("WordPress.com ‹ Log In").WebCheckBox("rememberme").Set "OFF"
			Browser("WordPress.com").Page("WordPress.com ‹ Log In").WebButton("Log In").Click

			On error resume next
			a = Browser("Following ‹ Reader — WordPress").Page("Following ‹ Reader — WordPress").WebElement("My Site").GetRoProperty("innertext")
			print a

			On error resume next
			b = Browser("Following ‹ Reader — WordPress").Page("WordPress.com ‹ Log In").WebElement("ERROR").GetRoProperty("innertext")
			print b
			'On error resume next

			if(a="My Site") Then
			MsgBox "Login Successful"

			Browser("WordPress.com").Page("Following ‹ Reader — WordPress").Link("Write").Click
			post =inputbox("Enter the title:")

			'''if(post=empty)Then 
			'''msgbox("Publish Button in Application will not be enabled")

			'''End If

			Browser("WordPress.com").Page("Edit Post ‹ Site Title").WebEdit("Edit title").Set post
			wait(5)

			Browser("WordPress.com").Page("Edit Post ‹ Site Title").WebButton("Publish").Click

			''validate publish button

			Browser("WordPress.com").Page("Edit Post ‹ Site Title").WebElement("View Post").Click

			wait(3)

			'Browser("WordPress.com").Page("Edit Post ‹ Site Title_3").Frame("wp-preview-68").

			'Browser("WordPress.com").Page("Edit Post ‹ Site Title_3").Frame("wp-preview-68").WebElement("content").GetRoProperty("innertext")

			'''c = Browser("WordPress.com").Page("Edit Post ‹ Site Title_4").Frame("wp-preview-76").Link("AutomationTest").GetRoProperty("innertext")
			wait 3
			c = Browser("Following ‹ Reader — WordPress").Page("Edit Post ‹ Site Title_2").Frame("wp-preview-226").Link("Content").GetRoProperty("innertext")

			'''If post=c Then
			'''MsgBox "Content Validated"
			'''End If
			
			if(post = c) then
			msgbox "Content Validated"
			c=empty
			post=empty
			else
			msgbox "Content is not there"
			c=empty
			post=empty
			end if

			Browser("WordPress.com").InsightObject("InsightObject").Click @@ hightlight id_;_4_;_script infofile_;_ZIP::ssf3.xml_;_

			
			Browser("WordPress.com").Page("Edit Post ‹ Site Title_2").Image("Me").Click

			Browser("WordPress.com").Page("My Profile — WordPress.com").WebButton("Sign Out").Click

			a=empty

			ElseIf (b = "ERROR") Then

			msgbox "Login Failed"
			Browser.navigate "https://wordpress.com/?appromo"
			b=empty

			end if

	next
			Set ExcelObj=nothing
			Set WorkbookObj=nothing
			set SheetObj=nothing


'End Of 1st Transaction

''Start Of 2nd Transaction

'''set ExcelObj = createobject("excel.application")
'''ExcelObj.visible = True
'''set WorkbookObj= ExcelObj.Workbooks.Open ("C:\Users\Balaji Vignesh\Desktop\Data_sheet1.xlsx")

'''set SheetObj = WorkbookObj.Worksheets("second innings")
'''Row=SheetObj.UsedRange.Rows.Count
'''For  i= 1 to Row

'''LoginUserId=SheetObj.cells(i,1).value

'''Password=SheetObj.cells(i,2).value
ElseIf sfid="SF2" Then
Browser.navigate "https://wordpress.com/?appromo"

Browser("WordPress.com").Page("WordPress.com").Link("Sign In").Click
Browser("WordPress.com").Page("WordPress.com ‹ Log In").WebEdit("log").Set "balajivigneshsm@gmail.com"
wait(1)
Browser("WordPress.com").Page("WordPress.com ‹ Log In").WebEdit("pwd").Set "16119192315184@bv"
wait(1)
Browser("WordPress.com").Page("WordPress.com ‹ Log In").WebCheckBox("rememberme").Set "OFF"
Browser("WordPress.com").Page("WordPress.com ‹ Log In").WebButton("Log In").Click

'''On error resume next

'''a = Browser("Following ‹ Reader — WordPress").Page("Following ‹ Reader — WordPress").WebElement("My Site").GetRoProperty("innertext")
'''On error resume next
'''b = Browser("Following ‹ Reader — WordPress").Page("WordPress.com ‹ Log In").WebElement("ERROR").GetRoProperty("innertext")
'On error resume next

'''if(a= "My Site") Then
'''MsgBox "Login Successful"

'''ElseIf ( b = "ERROR") Then
'''MsgBox "Login Failed"

'''End If 

Browser("Following ‹ Reader — WordPress").Page("Following ‹ Reader — WordPress").WebElement("My Site").Click
wait 2

Browser("Following ‹ Reader — WordPress").Page("Stats ‹ Site Title — WordPress").WebElement("Pages").Click
wait 2
Browser("Following ‹ Reader — WordPress").Page("Pages ‹ Site Title — WordPress").WebElement("Drafts").Click
wait 2
Browser("Following ‹ Reader — WordPress").Page("Pages ‹ Site Title — WordPress_2").Link("Start a Page").Click
wait 2
Browser("Following ‹ Reader — WordPress").Page("New Page ‹ Site Title").WebEdit("Edit title").Set "Kowshik"
wait 7
''Browser("Following ‹ Reader — WordPress").Page("Edit Page ‹ Site Title").Frame("Frame").WebElement("Content").set "Haram se"
'wait 7
''Browser("Following ‹ Reader — WordPress").Page("Edit Page ‹ Site Title").WebButton("Save").Click
'wait 2


Browser("Following ‹ Reader — WordPress").Page("Following ‹ Reader — WordPress").WebElement("My Site").Click
wait 2
Browser("Following ‹ Reader — WordPress").Page("Stats ‹ Site Title — WordPress").WebElement("Pages").Click
wait 2
Browser("Following ‹ Reader — WordPress").Page("Pages ‹ Site Title — WordPress").WebElement("Drafts").Click
wait 7
Browser("Following ‹ Reader — WordPress").InsightObject("InsightObject").Click
wait 7
Browser("Following ‹ Reader — WordPress").Page("Pages ‹ Site Title — WordPress_2").WebButton("Publish").Click
wait 2
Browser("Following ‹ Reader — WordPress").Page("Pages ‹ Site Title — WordPress_2").WebElement("Published").Click

'''if (v = Browser("Following ‹ Reader — WordPress").Page("Pages ‹ Site Title — WordPress_3").Link("Validate_Published").exist) then

'''msgbox "Drafted page was published"

'''endif

Browser("WordPress.com").Page("Edit Post ‹ Site Title_2").Image("Me").Click

Browser("WordPress.com").Page("My Profile — WordPress.com").WebButton("Sign Out").Click


'''Set ExcelObj=nothing
'''Set WorkbookObj=nothing
'''set SheetObj=nothing
End If
			End If
		End If
	Next
	End If 
Next

