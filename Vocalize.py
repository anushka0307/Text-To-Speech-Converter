import streamlit as st
import win32com.client as wincl
import pythoncom

pythoncom.CoInitialize()

sapi = wincl.Dispatch("SAPI.SpVoice")
voices = sapi.GetVoices()

def main():
    st.title("VOCALIZE")

    if st.button("Start"):
        st.session_state.started = True
        sapi.Speak("Welcome to VOCALIZE! Choose your voice.")

    if "started" not in st.session_state:
        st.session_state.started = False
        st.session_state.voice_selected = False
        st.session_state.text_spoken = False

    if st.session_state.started:
        
        voice_option = st.selectbox("Select Voice", ["Select", "Male", "Female"])

        if voice_option == "Male":
            sapi.Voice = voices.Item(0)
        elif voice_option == "Female":
            sapi.Voice = voices.Item(1)

        if voice_option != "Choose":
            st.session_state.voice_selected = True

        if st.session_state.voice_selected and (voice_option == "Male" or voice_option == "Female"):
            if not st.session_state.text_spoken:
                sapi.Speak("Enter your text.")
                st.session_state.text_spoken = True

            text_input = st.text_area("Enter text to speak")

            if st.button("Speak"):
                sapi.Speak(text_input)

if __name__ == "__main__":
    main()

pythoncom.CoUninitialize()
