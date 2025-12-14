import tkinter as tk
from tkinter import scrolledtext, messagebox
import speech_recognition as sr
import pyttsx3
import webbrowser
import datetime
import urllib.parse
import pywhatkit
import smtplib
import threading
import queue

# ---- Setup ----
recognizer = sr.Recognizer()
engine = pyttsx3.init()
voices = engine.getProperty('voices')

female_voice = next((v for v in voices if "female" in v.name.lower() or "zira" in v.name.lower()), None)
engine.setProperty('voice', female_voice.id if female_voice else voices[0].id)
engine.setProperty('rate', 180)

speak_lock = threading.Lock()
speaking_flag = threading.Event()  # Indicates if engine is currently speaking
tts_queue = queue.Queue()  # Queue for TTS messages


def tts_worker():
    """Worker thread for Text-to-Speech."""
    while True:
        text = tts_queue.get()
        if text is None:  # Sentinel value to stop the thread
            break
        try:
            with speak_lock:
                speaking_flag.set()  # Set flag when speaking starts
                engine.say(text)
                engine.runAndWait()
                speaking_flag.clear()  # Clear flag when speaking ends
        except Exception as e:
            print(f"TTS Error: {e}")
            log_output(f"Error speaking: {e}")
        finally:
            tts_queue.task_done()


tts_thread = threading.Thread(target=tts_worker, daemon=True)
tts_thread.start()


def stop_speaking():
    """Stops the current TTS output and clears the queue."""
    with speak_lock:
        if speaking_flag.is_set():
            engine.stop()
            speaking_flag.clear()
        # Clear any pending messages in the queue
        while not tts_queue.empty():
            try:
                tts_queue.get_nowait()
                tts_queue.task_done()
            except queue.Empty:
                break


def speak(text):
    """Adds text to the TTS queue to be spoken."""
    stop_speaking()  # Stop any currently speaking audio and clear queue
    log_output(f"Jarvis: {text}")
    tts_queue.put(text)


def log_output(message):
    """Logs messages to the GUI output box."""

    def inner():
        output_box.config(state=tk.NORMAL)
        output_box.insert(tk.END, message + '\n')
        output_box.config(state=tk.DISABLED)
        output_box.yview(tk.END)

    root.after(0, inner)  # Ensures GUI updates are on the main thread


def get_time():
    """Returns the current time."""
    now = datetime.datetime.now()
    return now.strftime("It is %I:%M %p on %A, %B %d, %Y.")  # More detailed time


def search_on_google(query):
    """Opens Google with the given query."""
    if not query:
        speak("Please tell me what to search for on Google.")
        return
    encoded = urllib.parse.quote_plus(query)
    search_url = f"https://www.google.com/search?q={encoded}"
    try:
        webbrowser.open(search_url)
        speak(f"Searching Google for {query}")
    except Exception as e:
        speak(f"Could not open browser for search. Error: {e}")
        log_output(f"Error opening browser: {e}")


def play_youtube_video(query):
    """Opens YouTube to play a video based on the query."""
    if not query:
        speak("Please tell me what video to play on YouTube.")
        return
    try:
        # pywhatkit.playonyt is more reliable for playing YouTube videos
        pywhatkit.playonyt(query)
        speak(f"Playing {query} on YouTube.")
    except Exception as e:
        speak(f"Could not play {query} on YouTube. Error: {e}")
        log_output(f"Error playing YouTube video: {e}")


def send_whatsapp_message():
    """Sends a WhatsApp message using pywhatkit."""
    phone = phone_entry.get().strip()
    message = whatsapp_msg_entry.get().strip()

    if not phone or not message:
        speak("Please enter both the recipient's phone number and the message for WhatsApp.")
        return

    # Basic validation for phone number format
    if not phone.startswith('+') or not phone[1:].isdigit():
        speak("Please enter a valid phone number, including the country code, like +911234567890.")
        return

    try:
        speak(
            "Attempting to send WhatsApp message. This may open your browser to WhatsApp Web. Please make sure you are logged in.")
        log_output(f"Sending WhatsApp to {phone}: {message}")

        now = datetime.datetime.now()
        send_hour = now.hour
        send_minute = now.minute + 1
        if send_minute >= 60: # Handle minute overflow to next hour
            send_minute = 0
            send_hour = (send_hour + 1) % 24

        pywhatkit.sendwhatmsg(phone_no=phone, message=message, time_hour=send_hour, time_min=send_minute, wait_time=20,
                              tab_close=True)
        speak("WhatsApp message scheduled and attempted to send.")
        log_output("WhatsApp message sent successfully (or window opened).")
    except pywhatkit.core.exceptions.CountryCodeException:
        speak("Invalid country code in phone number. Please use the full international format like +91.")
        log_output("WhatsApp Error: Invalid country code.")
    except Exception as e:
        speak("Failed to send WhatsApp message. Please ensure WhatsApp Web is logged in and the number is correct.")
        log_output(f"WhatsApp Error: {e}")


def send_email_message():
    """Sends an email using SMTPLib."""
    to_email = email_to_entry.get().strip()
    subject = email_subject_entry.get().strip()
    body = email_body_entry.get().strip()
    from_email = email_from_entry.get().strip()
    from_pass = email_pass_entry.get().strip()

    if not all([to_email, subject, body, from_email, from_pass]):
        speak("Please fill in all email fields: Your email, App Password, Recipient, Subject, and Body.")
        return

    try:
        speak("Attempting to send email.")
        log_output(f"Sending email from {from_email} to {to_email} with subject: {subject}")

        # Basic email format for message
        message = f"From: {from_email}\nTo: {to_email}\nSubject: {subject}\n\n{body}"

        server = smtplib.SMTP('smtp.gmail.com', 587) # Use Gmail's SMTP server and port
        server.starttls() # Enable TLS for security
        server.login(from_email, from_pass)
        server.sendmail(from_email, to_email, message)
        server.quit()
        speak("Email sent successfully.")
        log_output("Email sent successfully.")
    except smtplib.SMTPAuthenticationError:
        speak(
            "Email failed: Authentication error. For Gmail, use an App Password, not your regular password, especially if 2FA is on.")
        log_output("Email Error: SMTP Authentication failed. Check credentials/App Password.")
    except smtplib.SMTPConnectError as e:
        speak("Email failed: Could not connect to SMTP server. Check internet or server address.")
        log_output(f"Email Error: SMTP connection error: {e}")
    except smtplib.SMTPSenderRefused:
        speak("Email failed: Sender address refused. Check 'From Email' field.")
        log_output("Email Error: Sender address refused.")
    except Exception as e:
        speak("Email failed to send due to an unexpected error.")
        log_output(f"Email Error: {e}")


def handle_voice_command_thread():
    """Wrapper to run handle_voice in a separate thread."""
    threading.Thread(target=handle_voice, daemon=True).start()


def handle_voice():
    """Listens for voice commands and processes them."""
    stop_speaking()
    speak("Listening...")
    try:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source, duration=1)
            audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
            command = recognizer.recognize_google(audio, language="en-US").lower()
            log_output(f"You: {command}")

            if not command.strip():
                speak("I didn't hear anything. Please try again.")
                return

            if "time" in command:
                speak(get_time())
            elif "search google for" in command:
                query = command.replace("search google for", "").strip()
                search_on_google(query)
            elif "search" in command:
                query = command.replace("search", "").strip()
                if query:
                    speak("Do you want me to search Google or YouTube?")
                    with sr.Microphone() as source:
                        recognizer.adjust_for_ambient_noise(source, duration=0.5)
                        audio_choice = recognizer.listen(source, timeout=3, phrase_time_limit=3)
                        try:
                            choice = recognizer.recognize_google(audio_choice, language="en-US").lower()
                            if "google" in choice:
                                search_on_google(query)
                            elif "youtube" in choice:
                                play_youtube_video(query)
                            else:
                                speak("I didn't understand your choice. Searching Google by default.")
                                search_on_google(query)
                        except sr.UnknownValueError:
                            speak("Sorry, I didn't understand your choice. Searching Google by default.")
                            search_on_google(query)
                        except sr.WaitTimeoutError:
                            speak("No choice detected. Searching Google by default.")
                            search_on_google(query)
                else:
                    speak("What would you like to search for?")
            elif "play" in command and "youtube" in command:
                query = command.replace("play", "").replace("youtube", "").strip()
                play_youtube_video(query)
            elif "stop" in command or "exit" in command or "quit" in command:
                speak("Goodbye!")
                root.quit()
            else:
                speak(
                    "Command not recognized. You can ask me about the time, search Google or YouTube, or tell me to stop.")
    except sr.UnknownValueError:
        speak("Sorry, I didn't catch that. Please speak clearly.")
        log_output("Speech Recognition: Unknown value error.")
    except sr.WaitTimeoutError:
        speak("I didn't hear anything. Please try again.")
        log_output("Speech Recognition: Listening timed out (no speech detected).")
    except sr.RequestError as e:
        speak(
            f"Could not request results from Google Speech Recognition service; please check your internet connection.")
        log_output(f"Speech Recognition Request Error: {e}")
    except Exception as e:
        speak("An unexpected error occurred during voice command processing.")
        print(f"General Voice Handling Error: {e}")
        log_output(f"General Voice Handling Error: {e}")


# ---- GUI Setup ----
root = tk.Tk()
root.title("Jarvis Voice Assistant")
root.geometry("700x850")  # Increased height further
root.config(bg="#1e1e1e")

tk.Label(root, text="Jarvis - Voice Assistant", font=("Helvetica", 20, "bold"), fg="#0f62fe", bg="#1e1e1e").pack(
    pady=15)
output_box = scrolledtext.ScrolledText(root, width=80, height=15, state=tk.DISABLED, wrap=tk.WORD,
                                       font=("Courier New", 10), bg="#2c2c2c", fg="#00ff00", insertbackground="white")
output_box.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

# Main Voice Command Button
tk.Button(root, text="🎙️ Speak Command", font=("Arial", 16, "bold"), bg="#0f62fe", fg="white",
          command=handle_voice_command_thread,
          relief=tk.RAISED, bd=4).pack(pady=10)

# WhatsApp Section
whatsapp_frame = tk.LabelFrame(root, text="Send WhatsApp Message", bg="#1e1e1e", fg="white", font=("Arial", 12, "bold"),
                               padx=10, pady=10)
whatsapp_frame.pack(pady=10, padx=10, fill=tk.X)
tk.Label(whatsapp_frame, text="Phone No. (e.g., +911234567890):", fg="white", bg="#1e1e1e").pack(anchor='w')
phone_entry = tk.Entry(whatsapp_frame, width=50, bg="#3c3c3c", fg="white", insertbackground="white")
phone_entry.pack(pady=2, fill=tk.X, padx=5)
tk.Label(whatsapp_frame, text="Message:", fg="white", bg="#1e1e1e").pack(anchor='w')
whatsapp_msg_entry = tk.Entry(whatsapp_frame, width=50, bg="#3c3c3c", fg="white", insertbackground="white")
whatsapp_msg_entry.pack(pady=2, fill=tk.X, padx=5)
tk.Button(whatsapp_frame, text="Send WhatsApp",
          command=lambda: threading.Thread(target=send_whatsapp_message, daemon=True).start(), bg="#28a745",
          fg="white").pack(pady=5)

# Email Section
email_frame = tk.LabelFrame(root, text="Send Email", bg="#1e1e1e", fg="white", font=("Arial", 12, "bold"), padx=10,
                            pady=10)
email_frame.pack(pady=10, padx=10, fill=tk.X)
tk.Label(email_frame, text="From Email:", fg="white", bg="#1e1e1e").pack(anchor='w')
email_from_entry = tk.Entry(email_frame, width=50, bg="#3c3c3c", fg="white", insertbackground="white")
email_from_entry.pack(pady=2, fill=tk.X, padx=5)
tk.Label(email_frame, text="App Password (for Gmail):", fg="white", bg="#1e1e1e").pack(anchor='w')
email_pass_entry = tk.Entry(email_frame, show="*", width=50, bg="#3c3c3c", fg="white", insertbackground="white")
email_pass_entry.pack(pady=2, fill=tk.X, padx=5)
tk.Label(email_frame, text="To Email:", fg="white", bg="#1e1e1e").pack(anchor='w')
email_to_entry = tk.Entry(email_frame, width=50, bg="#3c3c3c", fg="white", insertbackground="white")
email_to_entry.pack(pady=2, fill=tk.X, padx=5)
tk.Label(email_frame, text="Subject:", fg="white", bg="#1e1e1e").pack(anchor='w')
email_subject_entry = tk.Entry(email_frame, width=50, bg="#3c3c3c", fg="white", insertbackground="white")
email_subject_entry.pack(pady=2, fill=tk.X, padx=5)
tk.Label(email_frame, text="Body:", fg="white", bg="#1e1e1e").pack(anchor='w')
email_body_entry = tk.Entry(email_frame, width=50, bg="#3c3c3c", fg="white", insertbackground="white")
email_body_entry.pack(pady=2, fill=tk.X, padx=5)
tk.Button(email_frame, text="Send Email",
          command=lambda: threading.Thread(target=send_email_message, daemon=True).start(), bg="#ffc107",
          fg="black").pack(pady=5)

# Ensure TTS thread is stopped on window close
def on_closing():
    print("Closing Jarvis...")
    tts_queue.put(None)  # Signal the TTS worker to stop
    tts_thread.join(timeout=2)  # Wait for the TTS thread to finish (with a timeout)
    root.destroy()


root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()


# Stop TTS worker when the GUI window is closed
tts_queue.put(None)
tts_thread.join()