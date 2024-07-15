import streamlit as st
import numpy as np
import random
import json
import pickle
import nltk
from nltk.stem import WordNetLemmatizer
from keras.models import load_model

# Download required NLTK data
nltk.download('punkt')
nltk.download('wordnet')
lemmatizer = WordNetLemmatizer()

# Load the model and data files with error handling
try:
    model = load_model(r'C:\Users\rouar\OneDrive\Bureau\multipleapp\pages\roua_best_chatbot_model.keras')
    with open(r'C:\Users\rouar\OneDrive\Bureau\multipleapp\roua_final_intents.json') as f:
        intents = json.load(f)
    with open(r'C:\Users\rouar\OneDrive\Bureau\multipleapp\pages\words.pkl', 'rb') as f:
        words = pickle.load(f)
    with open(r'C:\Users\rouar\OneDrive\Bureau\multipleapp\pages\classes.pkl', 'rb') as f:
        classes = pickle.load(f)
except Exception as e:
    st.error(f"Error loading files: {e}")
    st.stop()

# Prepare documents list for evaluation
documents = []
for intent in intents['intents']:
    for pattern in intent['patterns']:
        w = nltk.word_tokenize(pattern)
        documents.append((w, intent['tag']))

def clean_up_sentence(sentence):
    sentence_words = nltk.word_tokenize(sentence)
    sentence_words = [lemmatizer.lemmatize(word.lower()) for word in sentence_words]
    return sentence_words

def bow(sentence, words, show_details=True):
    sentence_words = clean_up_sentence(sentence)
    bag = [0] * len(words)
    for s in sentence_words:
        for i, w in enumerate(words):
            if w == s:
                bag[i] = 1
                if show_details:
                    st.write(f"found in bag: {w}")
    return np.array(bag)

def predict_class(sentence, model):
    p = bow(sentence, words, show_details=False)
    res = model.predict(np.array([p]))[0]
    ERROR_THRESHOLD = 0.25
    results = [[i, r] for i, r in enumerate(res) if r > ERROR_THRESHOLD]
    results.sort(key=lambda x: x[1], reverse=True)
    return_list = []
    for r in results:
        return_list.append({"intent": classes[r[0]], "probability": str(r[1])})
    return return_list

def getResponse(ints, intents_json):
    if not ints:
        return "I'm sorry, I don't understand that."
    
    tag = ints[0]['intent']
    list_of_intents = intents_json['intents']
    for i in list_of_intents:
        if i['tag'] == tag:
            result = random.choice(i['responses'])
            break
    return result

def chatbot_response(msg):
    ints = predict_class(msg, model)
    res = getResponse(ints, intents)
    return res

st.title("Chatbot")

# Initialize chat history
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []

# Handle user input
user_input = st.chat_input("You: ")

if user_input:
    response = chatbot_response(user_input)
    # Update chat history
    st.session_state.chat_history.append({"role": "user", "content": user_input})
    st.session_state.chat_history.append({"role": "assistant", "content": response})

# Display chat history
for message in st.session_state.chat_history:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])
