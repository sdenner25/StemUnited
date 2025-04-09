<%
' Set up the database connection
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")

' Connect to Access database (replace path if needed)
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("App_Data\Data.accdb")

' Get form values sent from the HTML page
Dim name, grade
name = Request.Form("student_name")
grade = Request.Form("student_grade")

' Build and run the SQL command
Dim sql
sql = "INSERT INTO Students (Student_Name, Student_Grade) VALUES ('" & name & "', '" & grade & "')"
conn.Execute(sql)

' Confirm for user
Response.Write("<h3>Thank you! The student was added to the database.</h3>")

' Clean up
conn.Close
Set conn = Nothing
%>