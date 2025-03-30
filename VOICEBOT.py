import sys
import time
import speech_recognition as sr
import webbrowser
import pyttsx3
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import subprocess

recognizer = sr.Recognizer()
paused = False
narration_active = False
engine = pyttsx3.init()
last_command = ""
pause_message_printed = False

def speak(text, stop_narration=False):
    global narration_active
    if stop_narration:
        if narration_active:
            engine.stop()
            narration_active = False
        return  # Exit early if stopping narration
    narration_active = True
    engine.say(text)
    engine.runAndWait()
    narration_active = False

def open_youtube_and_play(query):
    search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
    
    # Set up Selenium with Chrome options
    options = Options()
    options.binary_location = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"  # Update if necessary
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get(search_url)
    
    try:
        # Wait for the first video result to be clickable and then click it
        wait = WebDriverWait(driver, 10)
        first_video = wait.until(EC.element_to_be_clickable((By.ID, "video-title")))
        first_video.click()
        
        # Wait for the video to load
        time.sleep(7)  # Wait for the video to load

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        driver.quit()

def search_on_chrome(query):
    search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
    
    # Fetch the search results page
    response = requests.get(search_url)
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Extract snippets from the search results
    snippets = []
    for result in soup.find_all('div', class_='BNeawe s3v9rd AP7Wnd'):
        text = result.get_text()
        if text:
            # Break into sentences and select the first two
            sentences = text.split('. ')
            if len(sentences) >= 2:
                snippet = '. '.join(sentences[:2]) + '.'  # Join the first two sentences
            else:
                snippet = text
            snippets.append(snippet)
            if len(snippets) >= 1:  # Collect up to 1 snippet
                break
    
    # Create the summary from the collected snippets
    if snippets:
        summary = " ".join(snippets)
    else:
        summary = "No summary found."
    
    print(f"Summary: {summary}")
    speak(summary)  # Narration starts
    webbrowser.open(search_url)

def open_windows_app(app_name):
    app_paths = {
        "settings": "ms-settings:",
        "calculator": "calc",
        "mail": "mailto:",
        "spotify": "https://open.spotify.com"
    }

    if app_name in app_paths:
        try:
            if app_name == "spotify":
                webbrowser.open(app_paths[app_name])
                speak("Opening Spotify. You might need to sign in.")
            else:
                subprocess.Popen(app_paths[app_name], shell=True)
                speak(f"Opening {app_name}")
        except Exception as e:
            speak(f"Could not open {app_name}. {str(e)}")
    else:
        speak(f"Application {app_name} not recognized.")

def control_system_function(command):
    try:
        if command == "turn off wi-fi":
            subprocess.run(["netsh", "interface", "set", "interface", "Wi-Fi", "admin=disable"], check=True)
            speak("Wi-Fi turned off")
        elif command == "turn on wi-fi":
            subprocess.run(["netsh", "interface", "set", "interface", "Wi-Fi", "admin=enable"], check=True)
            speak("Wi-Fi turned on")
        elif command == "turn off bluetooth":
            subprocess.run(["powershell", "-Command", "Disable-Bluetooth"], check=True)
            speak("Bluetooth turned off")
        elif command == "turn on bluetooth":
            subprocess.run(["powershell", "-Command", "Enable-Bluetooth"], check=True)
            speak("Bluetooth turned on")
        elif command == "turn on airplane mode":
            subprocess.run(["powershell", "-Command", "Enable-AirplaneMode"], check=True)
            speak("Airplane mode turned on")
        elif command == "turn off airplane mode":
            subprocess.run(["powershell", "-Command", "Disable-AirplaneMode"], check=True)
            speak("Airplane mode turned off")
        elif command.startswith("change brightness to"):
            level = command.split("change brightness to ")[1]
            subprocess.run(["powershell", "-Command", f"(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,{level})"], check=True)
            speak(f"Brightness changed to {level} percent")
        elif command == "battery":
            battery_info = subprocess.run(["powershell", "-Command", "Get-WmiObject -Class Win32_Battery | Select-Object -ExpandProperty EstimatedChargeRemaining"], check=True, capture_output=True, text=True)
            battery_percentage = battery_info.stdout.strip()
            speak(f"Battery percentage is {battery_percentage} percent")
    except Exception as e:
        speak(f"Failed to execute system command: {command}. {str(e)}")

def process_voice_commands():
    global paused, last_command, pause_message_printed

    try:
        with sr.Microphone() as mic:
            recognizer.adjust_for_ambient_noise(mic, duration=0.2)
            print("Listening...")
            audio = recognizer.listen(mic, timeout=3, phrase_time_limit=5)  # Set timeout and phrase_time_limit for quick responses
            text = recognizer.recognize_google(audio)
            text = text.lower()
            print(f"Recognized: {text}")

            # Check if the command is the same as the last one
            if text == last_command:
                return  # Skip processing if the same command is detected
            
            last_command = text  # Update last_command with the new text

            if text == "open google":
                webbrowser.open("https://www.google.com")
            elif text == "open youtube":
                webbrowser.open("https://www.youtube.com")
            elif text.startswith("play"):
                search_query = text.replace("play ", "")
                open_youtube_and_play(search_query)
            elif text.startswith("search"):
                search_query = text.replace("search ", "")
                search_on_chrome(search_query)
            elif text in ["quit", "close", "terminate", "close python", "terminate python", "python terminate"]:
                speak("Terminating Python")
                print("Terminating Python...")
                sys.exit()  # Exit the program
            elif text in ["pause", "pause voice"]:
                paused = True
                pause_message_printed = False
                speak("Voice recognition paused. Say 'start' to resume.")
                print("Voice recognition paused. Say 'start' to resume.")
            elif text in ["start", "start voice"]:
                paused = False
                speak("Voice recognition resumed.")
                print("Voice recognition resumed.")
            elif text in ["stop"]:
                speak("", stop_narration=True)
                print("Narration stopped.")
            elif text.startswith("open"):
                app_name = text.replace("open ", "")
                open_windows_app(app_name)
            elif text.startswith("turn off") or text.startswith("turn on") or text.startswith("change brightness to") or text == "battery":
                control_system_function(text)
            else:
                speak("Command not recognized")
                print("Command not recognized")
    except sr.UnknownValueError:
        speak("Could not understand the audio")
        print("Could not understand the audio")
    except sr.WaitTimeoutError:
        speak("Listening timed out while waiting for phrase to start")
        print("Listening timed out while waiting for phrase to start")
    except sr.RequestError as e:
        speak(f"Could not request results from Google Speech Recognition service; {e}")
        print(f"Could not request results from Google Speech Recognition service; {e}")
    except Exception as e:
        speak(f"An unexpected error occurred; {e}")
        print(f"An unexpected error occurred; {e}")

while True:
    if not paused:
        try:
            process_voice_commands()
        except Exception as e:
            print(f"Error processing voice commands: {e}")
            # Brief sleep to prevent rapid retries and excessive CPU usage
            time.sleep(1)
    else:
        if not pause_message_printed:
            print("Voice recognition is paused. Say 'start' to resume.")
            pause_message_printed = True
        try:
            with sr.Microphone() as mic:
                recognizer.adjust_for_ambient_noise(mic, duration=0.2)
                audio = recognizer.listen(mic, timeout=3, phrase_time_limit=5)  # Set timeout and phrase_time_limit for quick responses
                text = recognizer.recognize_google(audio)
                text = text.lower()

                if text in ["start", "start voice"]:
                    paused = False
                    speak("Voice recognition resumed.")
                    print("Voice recognition resumed.")
                    pause_message_printed = False
        except sr.UnknownValueError:
            continue
        except sr.WaitTimeoutError:
            continue
        except sr.RequestError as e:
            continue
        except Exception as e:
            continue
