{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 3,
   "id": "7499ebee",
   "metadata": {},
   "outputs": [
    {
     "name": "stderr",
     "output_type": "stream",
     "text": [
      "2023-09-25 09:41:38.690 \n",
      "  \u001b[33m\u001b[1mWarning:\u001b[0m to view this Streamlit app on a browser, run it with the following\n",
      "  command:\n",
      "\n",
      "    streamlit run C:\\Users\\eyamashiro\\AppData\\Local\\anaconda3\\Lib\\site-packages\\ipykernel_launcher.py [ARGUMENTS]\n"
     ]
    }
   ],
   "source": [
    "import streamlit as st\n",
    "import pandas as pd\n",
    "import pyodbc\n",
    "\n",
    "# Función para cargar datos a SQL Server\n",
    "def insert_into_database(data, table_name):\n",
    "    # Configuración de la conexión\n",
    "    conn_str = (\n",
    "        r'DRIVER={ODBC Driver 17 for SQL Server};'\n",
    "        r'SERVER=SURDBP03;'\n",
    "        r'DATABASE=DesarrolloOLAP;'\n",
    "        r'Trusted_Connection=yes;'\n",
    "    )\n",
    "    conn = pyodbc.connect(conn_str)\n",
    "\n",
    "    # Insertar datos\n",
    "    cursor = conn.cursor()\n",
    "    \n",
    "    for index, row in data.iterrows():\n",
    "        placeholders = ', '.join(['?'] * len(row))\n",
    "        columns = ', '.join(row.keys())\n",
    "        sql = f\"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})\"\n",
    "        cursor.execute(sql, tuple(row))\n",
    "\n",
    "    conn.commit()\n",
    "    cursor.close()\n",
    "    conn.close()\n",
    "\n",
    "# Diseño de la aplicación Streamlit\n",
    "st.title(\"Importar datos de Excel a SQL Server\")\n",
    "\n",
    "uploaded_file = st.file_uploader(\"Elige un archivo Excel\", type=['xlsx'])\n",
    "\n",
    "if uploaded_file:\n",
    "    df = pd.read_excel(uploaded_file, engine='openpyxl')\n",
    "    st.write(df)\n",
    "\n",
    "    table_name = st.text_input(\"Nombre de la tabla en SQL Server:\", value=\"x_tabla_prueba\")\n",
    "\n",
    "    if st.button(\"Importar a SQL Server\"):\n",
    "        try:\n",
    "            insert_into_database(df, table_name)\n",
    "            st.success(\"Datos importados con éxito en la tabla \" + table_name)\n",
    "        except Exception as e:\n",
    "            st.error(f\"Ocurrió un error: {e}\")\n"
   ]
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.11.4"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
