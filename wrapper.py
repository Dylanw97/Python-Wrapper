class Wrapper:

    def __init__(self):

        # global standard libraries imports
        import os
        import re
        import tomllib
        import subprocess

        # global object declarations
        self.os = os
        self.re = re
        self.toml = tomllib
        self.subprocess = subprocess
        self.openpyxl = None
        self.smtp = None
        self.encoders = None
        self.MIMEBase = None
        self.MIMEText = None
        self.MIMEMultipart = None

        # global class variables
        self.current_working_directory = os.path.dirname(__file__)

        self.configuration_filename = "config.toml"
        self.configuration_file_path = os.path.join(self.current_working_directory, self.configuration_filename)

        self.log_filename = "log.log"
        self.log_file_path = os.path.join(self.current_working_directory, self.log_filename)

        self.current_environment_reference_name = ""

        # superclass configuration
        self.super_class_configuration = ""

        # subclass configuration locations
        self.export_configuration_file_location = ""
        self.export_configuration_text = ""

    def get_wrapper_super_configuration(self):

        configuration_string = \
            f"""
            test = ""
            """

        self.super_class_configuration = configuration_string

    def check_dependencies(self):

        if self.current_environment_reference_name == "email":

            try:

                import smtplib
                from email.mime.multipart import MIMEMultipart
                from email.mime.text import MIMEText
                from email.mime.base import MIMEBase
                from email import encoders

                self.smtp = smtplib
                self.MIMEMultipart = MIMEMultipart()
                self.MIMEText = MIMEText
                self.MIMEBase = MIMEBase
                self.encoders = encoders

            except ImportError:

                return False

        elif self.current_environment_reference_name == "excel":

            try:

                import openpyxl
                self.openpyxl = openpyxl

            except ImportError:

                return False

    def install_dependencies(self, dependencies):

        for package in dependencies:

            self.subprocess.check_call(['pip', 'install', package])

    def generate_new_configuration_file(self, configuration_name):

        if str(configuration_name).lower() == "email":

            self.export_configuration_file_location = self.os.path.join(
                self.current_working_directory,
                "email_config.toml"
            )

            self.export_configuration_text = Email.get_email_wrapper_configuration()

        elif str(configuration_name).lower() == "excel":
            pass
        elif str(configuration_name).lower() == "data":
            pass
        elif str(configuration_name).lower() == "logger":
            pass

        self.export_configuration_text = \
            self.re.sub(r'\n\s*', '\n', (
                self.re.sub(r'[^\S\n]+', ' ', self.export_configuration_text).strip())
            )

        with open(self.export_configuration_file_location, 'w') as file:

            file.write(self.export_configuration_text)

    def generate_append_for_configuration_file(self, configuration_name):

        if str(configuration_name).lower() == "email":

            self.export_configuration_file_location = self.os.path.join(
                self.current_working_directory,
                self.configuration_filename
            )

            self.export_configuration_text = Email.get_email_wrapper_configuration()

        elif str(configuration_name).lower() == "excel":
            pass
        elif str(configuration_name).lower() == "data":
            pass
        elif str(configuration_name).lower() == "logger":
            pass

        self.export_configuration_text = \
            self.re.sub(r'\n\s*', '\n', (
                self.re.sub(r'[^\S\n]+', ' ', self.export_configuration_text).strip())
            )

        with open(self.export_configuration_file_location, 'a') as file:

            file.write("\n")

            file.write(self.export_configuration_text)


class Email(Wrapper):

    def __init__(self, external_configuration=True, secure=True, attachment=False):

        # docstring
        """
        :param external_configuration: Makes the class search for a configuration files in the root directory
        :param secure: Attempts to send the email securely or not
        :param attachment: Whether there is an attachment specified
        """

        # retrieve super configuration
        super().__init__()

        self.email_wrapper_dependencies = []

        # retrieve configuration
        if external_configuration:

            try:

                with open(self.configuration_file_path, mode="rb") as fp:
                    self.email_wrapper_configuration = self.toml.load(fp)

            except FileNotFoundError as e:

                print(e)

        else:

            self.email_wrapper_configuration = self.toml.loads(self.get_email_wrapper_configuration())

        # set email criteria
        self.external_configuration_status = external_configuration
        self.secure_status = secure
        self.email_attachment_status = attachment

        # create object to hold email message
        self.email_message = None

        # Set up the SMTP server
        self.smtp_server = self.email_wrapper_configuration["email"]["smtp_server"]
        self.smtp_port = self.email_wrapper_configuration["email"]["smtp_port"]

        if self.secure_status:
            self.secure_smtp_port_1 = self.email_wrapper_configuration["email"]["secure_smtp_port_1"]
            self.secure_smtp_port_2 = self.email_wrapper_configuration["email"]["secure_smtp_port_2"]
            self.smtp_username = self.email_wrapper_configuration["email"]["smtp_username"]
            self.smtp_password = self.email_wrapper_configuration["email"]["smtp_password"]

        self.email_sender = self.email_wrapper_configuration["email"]["sender"]
        self.email_recipients = self.email_wrapper_configuration["email"]["recipients"]
        self.email_subject = self.email_wrapper_configuration["email"]["subject"]
        self.email_body = self.email_wrapper_configuration["email"]["body"]

        self.email_attachment_path = self.email_wrapper_configuration["email"]["attachment_path"]
        self.email_attachment_name = self.email_wrapper_configuration["email"]["attachment_name"]

    @staticmethod
    def get_email_wrapper_configuration():

        configuration_string = \
            f"""
            [email]
            smtp_server = "mobile.lincofood.com"
            smtp_port = 25
            secure_smtp_port_1 = 587    
            secure_smtp_port_2 = 465
            smtp_username = ""
            smtp_password = ""
            
            sender = "qad-etfr@baader.com"
            recipients = ["dylan.wisdom@baader.com"]
            subject = "Test Email"
            body = "Test Email"
            
            attachment_path = ""
            attachment_name = ""
            """

        return configuration_string

    def check_email_environment(self):

        self.current_environment_reference_name = "email"

        if not self.check_dependencies():
            self.install_dependencies(self.email_wrapper_dependencies)

    def build_email_headers(self):

        # set email headers
        self.email_message["From"] = self.email_sender
        self.email_message["To"] = ", ".join(self.email_recipients)
        self.email_message["Subject"] = self.email_subject

    def build_email_body(self):

        # Add body to email
        self.email_message.attach(self.MIMEText(self.email_body, 'plain'))

    def build_email_attachment(self):

        # Attach the file
        with open(self.email_attachment_path, 'rb') as attachment:
            part = self.MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        self.encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename = {self.email_attachment_name}')

        self.email_message.attach(part)

    def post_email_outbound(self):

        if self.secure_status:

            try:

                # Connect to the SMTP server
                with self.smtp.SMTP(self.smtp_server, self.secure_smtp_port_1) as server:
                    server.starttls()  # Start TLS encryption
                    server.login(self.smtp_username, self.smtp_password)  # Login to the SMTP server
                    server.send_message(self.email_message)  # Send the email

            except Exception as post_secure_email_error:

                print(post_secure_email_error)

                # Connect to the SMTP server
                with self.smtp.SMTP(self.smtp_server, self.secure_smtp_port_2) as server:
                    server.starttls()  # Start TLS encryption
                    server.login(self.smtp_username, self.smtp_password)  # Login to the SMTP server
                    server.send_message(self.email_message)  # Send the email

        else:

            # Connect to the SMTP server
            with self.smtp.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()  # Start TLS encryption
                server.send_message(self.email_message)  # Send the email

    def send_email(self):

        self.check_email_environment()

        # create object for email message
        self.email_message = self.MIMEMultipart

        self.build_email_headers()
        self.build_email_body()

        if self.email_attachment_status:
            self.build_email_attachment()

        self.post_email_outbound()


class Excel(Wrapper):

    def __init__(self, external_configuration=True):

        # docstring
        """

        """

        # retrieve super configuration
        super().__init__()

        self.excel_wrapper_dependencies = ["openpyxl"]

        # retrieve configuration
        if external_configuration:

            try:

                with open(self.configuration_file_path, mode="rb") as fp:
                    self.excel_wrapper_configuration = self.toml.load(fp)

            except FileNotFoundError as e:

                print(e)

        else:

            self.excel_wrapper_configuration = self.toml.loads(self.get_excel_wrapper_configuration())

        self.excel_workbook = None
        self.excel_active_worksheet = None
        self.excel_headers = []

    @staticmethod
    def get_excel_wrapper_configuration():
        configuration_string = \
            f"""
            [excel]
            """

        return configuration_string

    def check_excel_environment(self):

        self.current_environment_reference_name = "excel"

        if not self.check_dependencies():

            self.install_dependencies(self.excel_wrapper_dependencies)

    def create_excel_workbook(self):

        self.excel_workbook = self.Workbook()

    def select_active_excel_workbook(self):

        self.

    def iterate_data_for_excel(self):

        # Write data from the list of lists to the worksheet
        for row_index, row_data in enumerate(data_to_excel, start=1):
            for col_index, value in enumerate(row_data, start=1):
                ws.cell(row=row_index, column=col_index, value=value)

    def create_excel_file(self, data_to_excel):

        self.check_excel_environment()

        # Create a new workbook
        wb = Workbook()

        # Select the active worksheet
        ws = wb.active

        # Write data from the list of lists to the worksheet
        for row_index, row_data in enumerate(data_to_excel, start=1):
            for col_index, value in enumerate(row_data, start=1):
                ws.cell(row=row_index, column=col_index, value=value)

        # Save the workbook
        wb.save(cfg["excel"]["file"])



class Data(Wrapper):

    def __init__(self):

        # retrieve super configuration
        super().__init__()

        pass


class Logger(Wrapper):

    def __init__(self):
        # retrieve super configuration
        super().__init__()

        pass
