<%@ Language=VBScript %>
<%
	
	m_name = Request.Form ("Name")
	m_company = Request.form("Company")
	m_phone = Request.Form ("Phone")
	m_email = Request.Form ("Email")
	m_address = Request.form("Address")
	m_message = Request.form("Remarks")

	Dim MyBody
	Dim MyCDONTSMail
		
    	Set MyCDONTSMail = Server.CreateObject("CDONTS.NEWMAIL")
    	MyCDONTSMail.From= "ketan@extremesolutions.in"
        MyCDONTSMail.To= "kandla@sitaramtransport.com"
	MyCDONTSMail.Subject="Contact Form Submitted by " & " " & m_name
   
	MyBody = MyBody & "PERSONAL DETIALS" & " " & vbCrLf
	MyBody = MyBody & "" & " " & vbCrLf
    	MyBody = MyBody & "Full Name         :" & " " & m_name & " " & vbCrLf
	MyBody = MyBody & "Company           :" & " " & m_company & vbCrLf
	MyBody = MyBody & "Email             :" & " " & m_email & vbCrLf
    	MyBody = MyBody & "Phone No          :" & " " & m_phone & vbCrLf
	MyBody = MyBody & "Address           :" & " " & m_address & vbCrLf
	MyBody = MyBody & "Message           :" & " " & m_message & vbCrLf
	       
    	MyCDONTSMail.Body= MyBody
    	MyCDONTSMail.Send
    
    	set MyCDONTSMail=nothing
	Response.Redirect ("thanks.htm")
	

%>