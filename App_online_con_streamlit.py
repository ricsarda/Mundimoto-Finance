import streamlit as st
import subprocess

st.title("Mundimoto Finance")

st.write("Selecciona funcionalidad:")

if st.button("DAILY"):
    # Aquí llamas al script DAILY.py
    # Usamos subprocess para ejecutar el script de Python
    try:
        subprocess.run([
            "C:/Users/Ricardo Sarda/AppData/Local/Programs/Python/Python311/python.exe",
            "C:/Users/Ricardo Sarda/Desktop/Python/DAILY.py"
        ], check=True)
        st.success("DAILY.py se ejecutó correctamente.")
    except Exception as e:
        st.error(f"Error al ejecutar DAILY.py: {e}")

if st.button("Pólizas Credit Stock"):
    # Aquí llamas al script Credit stock.py
    try:
        subprocess.run([
            "C:/Users/Ricardo Sarda/AppData/Local/Programs/Python/Python311/python.exe",
            "C:/Users/Ricardo Sarda/Desktop/Python/Credit stock.py"
        ], check=True)
        st.success("Credit stock.py se ejecutó correctamente.")
    except Exception as e:
        st.error(f"Error al ejecutar Credit stock.py: {e}")
