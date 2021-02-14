'Gmail config...

Set objMail = CreateObject("CDO.Message")
Set objConf = CreateObject("CDO.Configuration")
Set objFlds = objConf.Fields
objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'cdoSendUsingPort
objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com" 'your smtp server domain or IP address goes here
objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 'default port for email
'uncomment next three lines if you need to use SMTP Authorization
objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 20
objFlds.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic
objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "From-email" '(Gmail)
objFlds.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "password"

objFlds.Update
objMail.Configuration = objConf
objMail.From = "From-email"
objMail.To = "To-email"
objMail.Subject = "Email Subject Text"
objMail.TextBody = "The message of the email..."
objMail.AddAttachment "C:\Users\User\Desktop\Folder\emailv2.vbs"
objMail.Send
Set objFlds = Nothing
Set objConf = Nothing
Set objMail = Nothing
