import streamlit as st
from auth import check_auth
from main_view import show_main

def main():
    if check_auth():
        show_main()

if __name__ == "__main__":
    main()
