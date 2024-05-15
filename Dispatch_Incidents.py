import tkinter as tk
from tkinter import ttk, messagebox
import xml.etree.ElementTree as ET
from datetime import datetime
import os

def get_xml_filename():
    directory = "c:/swdtools/"
    if not os.path.exists(directory):
        os.makedirs(directory)
    return f"{directory}incident_dispatch_{datetime.now().strftime('%Y-%m')}.xml"

def init_xml():
    xml_file = get_xml_filename()
    if not os.path.exists(xml_file):
        root = ET.Element("Enregistrements")
        tree = ET.ElementTree(root)
        tree.write(xml_file)
    else:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        ensure_ids(root)
        tree.write(xml_file)
    return xml_file

def ensure_ids(root):
    id_counter = 1
    for record in root.findall('Enregistrement'):
        if 'id' not in record.attrib:
            record.set('id', str(id_counter))
        id_counter += 1

def save_incident(incident, agent, xml_file):
    if not incident or "INC" not in incident.upper():
        messagebox.showerror("Erreur", "Le champ 'Incident' doit contenir 'INC' et ne peut être vide.")
        return
    
    tree = ET.parse(xml_file)
    root = tree.getroot()
    today = datetime.now().strftime("%Y-%m-%d")
    
    existing_incident = next((rec for rec in root.findall('Enregistrement') 
                              if rec.find('Incident').text == incident and 
                                 rec.find('DateTime').text.startswith(today)), None)
    if existing_incident is not None:
        messagebox.showwarning("Attention", "Un ticket identique a déjà été enregistré aujourd'hui.")
        return

    max_id = max((int(rec.attrib['id']) for rec in root.findall('Enregistrement')), default=0)
    record = ET.Element("Enregistrement", id=str(max_id + 1))
    ET.SubElement(record, "Incident").text = incident
    ET.SubElement(record, "Agent").text = agent
    ET.SubElement(record, "DateTime").text = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    root.append(record)
    
    tree.write(xml_file)
    display_records(xml_file)
    display_agent_stats(xml_file)

def display_records(xml_file):
    today = datetime.now().strftime("%Y-%m-%d")
    tree_view = ET.parse(xml_file)
    root = tree_view.getroot()
    for i in tree.get_children():
        tree.delete(i)

    for record in root.findall('Enregistrement'):
        if record.find('DateTime').text.startswith(today):
            incident = record.find('Incident').text
            agent = record.find('Agent').text
            date_time = record.find('DateTime').text
            tree.insert("", tk.END, iid=record.attrib['id'], values=(incident, agent, date_time))

def delete_record(xml_file):
    if not incident_id.get():
        messagebox.showerror("Erreur", "Aucun enregistrement sélectionné")
        return
    tree_view = ET.parse(xml_file)
    root = tree_view.getroot()
    for record in root.findall('Enregistrement'):
        if record.attrib['id'] == incident_id.get():
            root.remove(record)
            break
    tree_view.write(xml_file)
    display_records(xml_file)
    display_agent_stats(xml_file)
    incident_id.set("")

def display_agent_stats(xml_file):
    tree_view = ET.parse(xml_file)
    root = tree_view.getroot()
    agent_stats_today = {}
    agent_stats_month = {}

    today = datetime.now().strftime("%Y-%m-%d")
    month = datetime.now().strftime("%Y-%m")

    for record in root.findall('Enregistrement'):
        agent = record.find('Agent').text
        date_time = record.find('DateTime').text
        if date_time.startswith(today):
            if agent in agent_stats_today:
                agent_stats_today[agent] += 1
            else:
                agent_stats_today[agent] = 1
        if date_time.startswith(month):
            if agent in agent_stats_month:
                agent_stats_month[agent] += 1
            else:
                agent_stats_month[agent] = 1

    stats_text_today = "Tickets d'aujourd'hui par agent:\n" + "\n".join(f"{agent}: {count}" for agent, count in agent_stats_today.items())
    stats_text_month = "Tickets du mois par agent:\n" + "\n".join(f"{agent}: {count}" for agent, count in agent_stats_month.items())
    stats_text = stats_text_today + "\n\n" + stats_text_month
    lbl_agent_stats.config(text=stats_text)

def open_info_window():
    info_window = tk.Toplevel(window)
    info_window.title("À propos")
    tk.Label(info_window, text="Auteur: Jonathan Cassara").pack(pady=10)
    tk.Label(info_window, text="Version: 1 Bêta 15/05/2024").pack(pady=10)
    info_window.geometry("200x100")

def select_item(event):
    selected = tree.selection()
    if selected:
        selected_id = selected[0]
        incident_id.set(selected_id)
        item_values = tree.item(selected_id, 'values')
        entry_incident.delete(0, tk.END)
        entry_incident.insert(0, item_values[0])
        combo_agent.set(item_values[1])

def init_gui(xml_file):
    global window
    window = tk.Tk()
    window.title("Gestion des Incidents")

    global lbl_agent_stats
    lbl_agent_stats = tk.Label(window, text="")
    lbl_agent_stats.grid(column=0, row=0, columnspan=2, sticky="w")

    btn_info = tk.Button(window, text="i", command=open_info_window)
    btn_info.grid(column=2, row=0, sticky=tk.W)

    lbl_incident = tk.Label(window, text="Incident:")
    lbl_incident.grid(column=0, row=1)

    global entry_incident
    entry_incident = tk.Entry(window)
    entry_incident.grid(column=1, row=1)

    lbl_agent = tk.Label(window, text="Agent:")
    lbl_agent.grid(column=0, row=2)

    global combo_agent
    combo_agent = ttk.Combobox(window, values=["Agent 1", "Agent 2", "Agent 3", "Agent 4"], state="readonly")
    combo_agent.grid(column=1, row=2)

    btn_save = tk.Button(window, text="Enregistrer", command=lambda: save_incident(entry_incident.get(), combo_agent.get(), xml_file))
    btn_save.grid(column=1, row=3)

    btn_delete = tk.Button(window, text="Supprimer", command=lambda: delete_record(xml_file))
    btn_delete.grid(column=1, row=4)

    global tree
    tree = ttk.Treeview(window, columns=("Incident", "Agent", "DateTime"), show="headings")
    tree.heading("Incident", text="Incident")
    tree.heading("Agent", text="Agent")
    tree.heading("DateTime", text="Date et Heure")
    tree.grid(column=0, row=5, columnspan=3)

    tree.bind('<ButtonRelease-1>', select_item)

    global incident_id
    incident_id = tk.StringVar()

    display_records(xml_file)
    display_agent_stats(xml_file)

    window.mainloop()

xml_file = init_xml()
init_gui(xml_file)