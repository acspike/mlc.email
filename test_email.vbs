msgbox "testing"
Dim Email
Set Email = CreateObject("MLC.Email.1")
With Email
    .setServer "mail.example.com"
    .setFrom "outbox@example.com"
    .setSubject "Test Message"
    .addTo "inbox@example.com"
    .addBcc "blind@example.com"
    .setHeader "X-MLC-Test", "Just Testing"
    .addText "This is a test message from VBS with an attachment"
    .addFile "test_email.vbs"
    .send
End With
Set Email = Nothing
