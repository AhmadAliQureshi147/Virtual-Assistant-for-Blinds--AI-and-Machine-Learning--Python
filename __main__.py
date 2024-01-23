import json
import speech_recognition as sr
import win32com.client
import os
import webbrowser
import wikipedia
import requests
import time
import re
import PyPDF2
import winsound
import datetime
import cv2
import torch
import pyttsx3
from datetime import datetime
from googlesearch import search
SETTINGS_FILE = 'settings.json'

def updating_the_settings(key, value):
    try:
        with open(SETTINGS_FILE, 'r') as f:
            settings = json.load(f)
    except FileNotFoundError:
        settings = {}

    settings[key] = value

    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f)

# Function to get settings
def get_settings(key, default_value):
    try:
        with open(SETTINGS_FILE, 'r') as f:
            settings = json.load(f)
            return settings.get(key, default_value)
    except FileNotFoundError:
        return default_value

def vocalize(text, sound_signal=None):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    # Get the rate from settings, default to 0 (normal speed)
    rate = get_settings('voice_rate', 0)
    speaker.Rate = rate

    speaker.Speak(text)
    if sound_signal:
        winsound.Beep(sound_signal, 500)

def play_welcome_sound():
    winsound.Beep(1000, 500)
    vocalize("Your assistant is now ready.", 2000)


def alter_voice_tempo():
    vocalize("Interested in modifying speech pace? Respond with either yes or no.")
    user_feedback = capture_user_input()

    if user_feedback:
        if "yes" in user_feedback.lower():
            vocalize("Please indicate your preference. Choose slow for a reduced pace, fast for a quicker pace, or stick with normal.")
            user_preference = capture_user_input()

            if user_preference:
                if "slow" in user_preference.lower():
                    updating_the_settings('pace_of_speech', -10)
                elif "fast" in user_preference.lower():
                    updating_the_settings('pace_of_speech', 10)
                elif "normal" in user_preference.lower():
                    updating_the_settings('pace_of_speech', 0)
                else:
                    vocalize("Couldn't comprehend your selection. Maintaining the current pace.")
            else:
                vocalize("Sorry, couldn't catch that. No changes will be made to the speech pace.")
        else:
            vocalize("Okay, the current pace will remain unchanged.")
    else:
        vocalize("Didn't understand your input. Retaining the existing speech pace.")


def perform_object_identification():
    speech_engine = pyttsx3.init()

    obj_model = torch.hub.load('ultralytics/yolov5', 'yolov5l', trust_repo=True)

    video_cap = cv2.VideoCapture(0)

    while True:
        # Fetch frames from the webcam
        status, image_frame = video_cap.read()
        if not status:
            print("Could not fetch frame.")
            break

        object_results = obj_model(image_frame)
        object_results.render()

        dataframe = object_results.pandas().xyxy[0]

        detected_labels = ""

        for index in dataframe.index:
            min_x, min_y = int(dataframe['xmin'][index]), int(dataframe['ymin'][index])
            max_x, max_y = int(dataframe['xmax'][index]), int(dataframe['ymax'][index])
            obj_label = dataframe['name'][index]
            detected_labels += obj_label + ", "

            cv2.rectangle(image_frame, (min_x, min_y), (max_x, max_y), (0, 255, 255), 2)
            cv2.putText(image_frame, obj_label, (min_x, min_y), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 255), 2)

        if detected_labels:
            speech_engine.say(f"Identified {detected_labels}")
            speech_engine.runAndWait()

        cv2.imshow('Real-time Object Identification', image_frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    video_cap.release()
    cv2.destroyAllWindows()

def capture_user_input():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            print("Recognizing ...")
            query = r.recognize_google(audio, language="en")
            print(f"User said: {query}")
            winsound.Beep(1000, 200)  # Command acknowledgment beep
            return query
        except Exception as e:
            winsound.Beep(500, 500)  # Error beep
            print("Sorry, could not recognize your voice.")
            return None


def initiate_navigation(location):
    try:
        webbrowser.open(f"https://www.google.com/maps/search/{location}")
        vocalize(f"Navigating to {location}")
    except Exception as e:
        vocalize(f"An error occurred: {e}")



def query_wikipedia(query):
    try:
        summary_data = wikipedia.summary(query, sentences=2)
        return summary_data
    except wikipedia.exceptions.DisambiguationError:
        winsound.Beep(500, 500)  # Alert sound
        return "Your search term is ambiguous. Could you provide more details?"
    except wikipedia.exceptions.PageError:
        winsound.Beep(500, 500)  # Alert sound
        return "The search term doesn't correspond to any Wikipedia articles."
    except Exception as unknown_error:
        winsound.Beep(500, 500)  # Alert sound
        return f"Encountered an unknown error: {unknown_error}"


def search_google(query):
    try:
        for j in search(query, num_results=1):
            return j
    except Exception as e:
        winsound.Beep(500, 500)  # Error beep
        return f"An error occurred during the Google search: {e}"

def handling_the_feedback(feedback_text, feedback_type, priority):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("user_feedback.txt", "a") as f:
        f.write(f"Timestamp: {timestamp}, Type: {feedback_type}, Priority: {priority}, Feedback: {feedback_text}\n")


def announce_day_and_date():
    now = datetime.now()
    day = now.strftime("%A")
    date_str = now.strftime("%Y-%m-%d")
    vocalize(f"Today is {day}, {date_str}")

def fetch_news_headlines():
    API_KEY = ""  # Replace with your own actual News API key
    BASE_URL = "https://newsapi.org/v2/top-headlines?country=us&apiKey=" + API_KEY
    response = requests.get(BASE_URL)
    news_data = response.json()

    if news_data['status'] == 'ok':
        articles = news_data['articles'][:5]
        for i, article in enumerate(articles):
            title = article['title']
            description = article['description']
            vocalize(f"News {i + 1}: {title}. {description}")
    else:
        vocalize("Sorry, I can't fetch the news right now.")


def create_reminder():
    vocalize("What do you want to be reminded about?")
    reminder = capture_user_input()
    print(f"Debug: Reminder is set to {reminder}")
    if reminder:
        vocalize("In how many minutes?")
        time_in_minutes = capture_user_input()
        print(f"Debug: Time in minutes is set to {time_in_minutes}")
        if time_in_minutes:
            try:

                time_in_minutes = re.findall(r'\d+', time_in_minutes)[0]
                time_in_seconds = float(time_in_minutes)
                print(f"Debug: Time in seconds is set to {time_in_seconds}")
                time.sleep(time_in_seconds)
                vocalize(f"Reminder: {reminder}")
            except (ValueError, IndexError) as e:
                vocalize("Could not understand the time input. Please try setting the reminder again.")
                print(f"Debug: Error is {e}")
        else:
            vocalize("Sorry, I couldn't capture the time. Please try again.")
    else:
        vocalize("Sorry, I couldn't capture what you want to be reminded about. Please try again.")


def read_from_text():
    try:
        path_to_file = "D:\\Downloads\\challenge.txt"  # Modify with the actual path to your text file
        with open(path_to_file, 'r') as file_handle:
            content = file_handle.read()
            vocalize(content)
    except Exception as error:
        vocalize(f"Failed to read the file due to: {error}")
        print(f"Debug info: Failed to read the file due to: {error}")



def read_and_speak_pdf():
    try:
        pdf_location = "D:\\Downloads\\Assignment.pdf"  # Adjust this to your actual file location
        with open(pdf_location, 'rb') as pdf_file:
            pdf_processor = PyPDF2.PdfReader(pdf_file)
            total_pages = len(pdf_processor.pages)

            for pageIndex in range(total_pages):
                currentPage = pdf_processor.pages[pageIndex]
                text = currentPage.extract_text()
                vocalize(text)

    except Exception as errorMsg:
        vocalize(f"A problem occurred while processing the PDF: {errorMsg}")
        print(f"Debug Info: Issue encountered during PDF processing: {errorMsg}")


def handling_feedback(feedback_text, feedback_type, priority):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("user_feedback.txt", "a") as f:
        f.write(f"Timestamp: {timestamp}, Type: {feedback_type}, Priority: {priority}, Feedback: {feedback_text}\n")
def collect_feedback(initial_command=None):
    if initial_command and "feedback" in initial_command.lower():
        vocalize("It seems you're interested in providing feedback. Would you like to proceed?")
    else:
        vocalize("Would you like to provide some feedback?")

    feedback = capture_user_input()

    if feedback and "yes" in feedback.lower():
        vocalize("What type of feedback is this? You can say 'functionality', 'speed', or 'error'")
        feedback_type = capture_user_input()

        vocalize("What is the priority level? You can say 'high', 'medium', or 'low'")
        priority = capture_user_input()

        vocalize("Please provide your detailed feedback.")
        feedback_text = capture_user_input()

        if feedback_text and feedback_type and priority:
            handling_the_feedback(feedback_text, feedback_type, priority)
            vocalize("Thank you for your feedback.")
        else:
            vocalize("Sorry, I couldn't capture your feedback.")

    elif feedback and "no" in feedback.lower():
        vocalize("Okay, no problem. If you have feedback in the future, feel free to share.")
    else:
        vocalize("Sorry, I didn't get that. Moving on.")

def fetch_weather_details():
    API_KEY = ""  # Replace with your own API key from OpenWeatherMap
    BASE_URL = "http://api.openweathermap.org/data/2.5/weather?"
    vocalize("Please tell me the city you want the weather for.")
    city = capture_user_input()
    if city:
        complete_api_link = f"{BASE_URL}q={city}&appid={API_KEY}"
        api_link = requests.get(complete_api_link)
        api_data = api_link.json()

        if api_data['cod'] == '404':
            vocalize("Invalid city: Please check your city name")
        else:
            main = api_data['main']
            wind = api_data['wind']
            weather_desc = api_data['weather'][0]['description']
            temp = round((main['temp'] - 273.15), 2)
            humidity = main['humidity']
            wind_speed = wind['speed']
            vocalize(f"Weather in {city}: {weather_desc}, Temperature: {temp}°C, Humidity: {humidity}%, Wind Speed: {wind_speed} m/s")

if __name__ == '__main__':
    play_welcome_sound()
    vocalize("Hello, I'm your Digital Assistant, Jarvis! How may I assist you today?")

    while True:
        print("Listening...")
        userInput = capture_user_input()

        if userInput:

            if "modify voice pace" in userInput.lower():
                alter_voice_tempo()

            if "tell me about" in userInput:
                topic = userInput.split("about")[-1].strip()
                description = query_wikipedia(topic)

                if "No Wikipedia page found for the query" in description:
                    google_info = search_google(topic)
                    description = f"No Wikipedia entry found. You might find this helpful: {google_info}"

                vocalize(f"Here's what I discovered about {topic}: {description}")

            for platform in [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.org"], ["google", "https://www.google.com"]]:
                if f"launch {platform[0]}" in userInput.lower():
                    webbrowser.open(platform[1])
                    vocalize(f"Launching {platform[0]}.")

            if "play the music" in userInput.lower():
                tune_file = "D:\\Downloads\\downfall-21371.mp3"  # Replace with your own file path
                os.startfile(tune_file)

            if "current time" in userInput.lower():
                currentTime = datetime.datetime.now().strftime("%I:%M:%S %p")
                vocalize(f"The current time is {currentTime}")

            if "how's the weather" in userInput.lower():
                fetch_weather_details()

            if "provide feedback" in userInput.lower():
                collect_feedback(userInput)

            if "reminder" in userInput.lower():
                create_reminder()

            if "breaking news" in userInput.lower():
                fetch_news_headlines()

            if "text file" in userInput.lower():
                read_from_text()

            if "pdf file" in userInput.lower():
                read_and_speak_pdf()

            if "day is it" in userInput.lower() or "current date" in userInput.lower():
                announce_day_and_date()

            if "how do you feel" in userInput.lower():
                vocalize("I'm functioning optimally! How may I assist you further?")

            if "find location" in userInput.lower():
                vocalize("Sure, what's the destination?")
                target_location = capture_user_input()
                if target_location:
                    initiate_navigation(target_location)

            if "identify objects" in userInput.lower():
                vocalize("Starting object identification. Press 'q' to stop.")
                perform_object_identification()

            if "terminate" in userInput.lower():
                vocalize("Goodbye! Take care!")
                break
