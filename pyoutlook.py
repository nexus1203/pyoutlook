import win32com.client

class Outlook:
    def __init__(self,):
        self.olMailItem = 0x0
        self.outlook_ = win32com.client.Dispatch("Outlook.Application")
        
    def create_email(self, subject:str, to:str):
        """create an email with subject and to

        Args:
            subject (str): subject of email
            to (str): mail address of receiver
            cc (str, optional): mail address of cc . Defaults to None.
        """
        self.mail = self.outlook_.CreateItem(self.olMailItem)
        self.mail.Subject = subject
        self.mail.To = to
    
    def attach_file(self, file_path:str):
        """Attach file to email.

        Args:
            file_path (str): path of file
        """
        self.mail.Attachments.Add(file_path)
    
    def send_email(self, body:str):
        """Send email with body text.

        Args:
            body (str): body of email
        """
        self.mail.Body = body
        self.mail.Send()
    
    def send_html_email(self, body:str):
        """ 
        Send email with html body text.

        Args:
            body (str): body of email
        """
        self.mail.HTMLBody = body
        self.mail.Send()
    
    def close_session(self):
        """Close outlook session."""
        self.outlook_.Quit()
        
