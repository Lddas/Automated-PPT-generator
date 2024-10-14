import tkinter as tk
from tkinter import filedialog, ttk
from pptx import Presentation
import json
import copying_and_modifying_slide
import day_plan
from pptx.util import Inches
import importlib
import inputs  # Import the inputs.py file initially
import os

script_dir = os.path.dirname(__file__)

def print_link(link, language=None):
    if language is not None:
        link = f"slides/{language}/" + link
    else:
        link = f"slides/" + link
    abs_file_path = os.path.join(script_dir, link)
    return abs_file_path

# Import data from the index.json file
with open(print_link("index.json")) as f:
    data = json.load(f)

# Initialize lists to store all the hotel dropdowns and room number entries
hotel_vars = []
room_number_vars = []

# Create a list to store day tabs and their step selections
day_frames = []
day_selections = []
day_dates = []



def extract_interface_data():
    # Extracting data from the Tkinter interface (client, dates, etc.)
    form_data = {
        "client": client_entry.get(),
        "type": event_type_entry.get(),
        "logo_path": logo_entry.get(),
        "num_days": days_entry.get(),
        "num_nights": nights_entry.get(),
        "num_people": people_entry.get(),
        "dates": dates_entry.get(),
        "selected_hotels": [],  # We'll add the selected hotels 
        "english_version": english_version_var.get()  # Add the boolean value
    }

    # Collect all selected hotels and their corresponding number of rooms
    for i, hotel_var in enumerate(hotel_vars):
        hotel_name = hotel_var.get()
        room_number = room_number_vars[i].get()
        if hotel_name != "Sélectionnez un hôtel":  # Only keep valid selections
            form_data["selected_hotels"].append({
                "hotel_name": hotel_name,
                "room_number": room_number
            })

    return form_data



def fill_input_file(form_data):
    # Prepare the data to be written in JSON format
    input_data = {
        "input_0_A": form_data["logo_path"],
        "input_0_B": form_data["dates"],
        "input_0_C": form_data["type"],
        "input_1_A": form_data["client"],
        "input_1_B": form_data["num_people"],
        "input_1_C": form_data["num_days"],
        "input_1_D": form_data["num_nights"],
        "input_1_E": f'Du {form_data["dates"]}',
        "input_multiple_hotels": len(form_data["selected_hotels"]) > 1,
    }

    # Handle selected hotels and their room numbers
    for idx, hotel_info in enumerate(form_data["selected_hotels"], start=12):
        hotel_name = hotel_info["hotel_name"]
        room_number = hotel_info["room_number"]

        # Example: Map hotels to specific input numbers based on their name
        if "KOUTOUBIA" in hotel_name:
            input_data["input_koutoubia_5"] = room_number
        elif "SOFITEL" in hotel_name:
            input_data["input_sofitel_3"] = room_number
        else:
            input_data[f'input_{idx}'] = room_number

    # Write the data to a JSON file
    with open("inputs.json", "w") as f:
        json.dump(input_data, f)



# Function to generate PowerPoint by copying slides based on selections
def generate_ppt():
    # Create an empty presentation (output presentation)
    outputPres = Presentation()
    outputPres.slide_width = Inches(13.33)
    outputPres.slide_height = Inches(7.5)

    # Extract form data from the interface
    form_data = extract_interface_data()
    fill_input_file(form_data)  # Fill the JSON file with new data

    # Load the updated inputs from the JSON file immediately after updating
    with open("inputs.json", "r") as f:
        inputs = json.load(f)  # Ensure inputs are loaded after modifying JSON

    # First slides (testing without images for now)
    for i in range(1, 7):
        try:
            if form_data["english_version"] == True:
                ppt = print_link("Slides_DEBUT_FIN_ANGLAIS.pptx", "ANGLAIS")
            else:
                ppt = print_link("Slides_DEBUT_FIN_FRANCAIS.pptx", "FRANCAIS")
            copying_and_modifying_slide.NewSlide(i, ppt, outputPres)
        except Exception as e:
            print(f"Error adding slide {i}: {e}")

    if len(form_data["selected_hotels"]) > 1:
        try:
            if form_data["english_version"] == True:
                ppt = print_link("liste_HOTELS_ANGLAIS.pptx", "ANGLAIS")
            else:
                ppt = print_link("liste_HOTELS_FRANCAIS.pptx", "FRANCAIS")
            copying_and_modifying_slide.NewSlide(1, ppt, outputPres)
        except Exception as e:
            print(f"Error adding hotel slide: {e}")

    # Adding hotel slides
    for hotel_info in form_data["selected_hotels"]:
        hotel_name = hotel_info["hotel_name"]

        link = data[hotel_name]["ppt"]
        slide_index = data[hotel_name]["index"]
        nb_slides = data[hotel_name]["nb_slides"]
        
        # Copy relevant slides
        for i in range(slide_index - 1, slide_index + nb_slides - 1):
            try:
                if form_data["english_version"] == True:
                    ppt = print_link(link.replace(".pptx", "_ANGLAIS.pptx"), "ANGLAIS")
                else:
                    ppt = print_link(link.replace(".pptx", "_FRANCAIS.pptx"), "FRANCAIS")
                copying_and_modifying_slide.NewSlide(i, ppt, outputPres)
            except Exception as e:
                print(f"Error adding activity slide {i}: {e}")

    # ADDING AGENDA
    copying_and_modifying_slide.make_agenda(outputPres, day_selections, form_data["english_version"])

    # Loop over each day and its selections (activities)
    for day_index, day_plan_info in enumerate(day_selections[:-1]):
        day_date = day_dates[day_index].get()  # Get the date for this day

        # Slides entre chaque jours
        if len(day_plan_info) >= 2 and len(day_plan_info) < 7:
            copying_and_modifying_slide.CopyAndModifySlide(len(day_plan_info) - 1, outputPres, day_plan_info, day_index + 1, day_date, form_data["english_version"])

        # Add other conditions for different lengths of day_plan_info...

        # For each selected activity in a day
        for etape in day_plan_info:
            link = data[etape]["ppt"]
            if link == "FALSE":
                continue
            if form_data["english_version"] == True:
                ppt = print_link(link.replace(".pptx", "_ANGLAIS.pptx"), "ANGLAIS")
            else:
                ppt = print_link(link.replace(".pptx", "_FRANCAIS.pptx"), "FRANCAIS")
            slide_index = data[etape]["index"]
            nb_slides = data[etape]["nb_slides"]
            
            # Copy relevant slides
            for i in range(slide_index - 1, slide_index + nb_slides - 1):
                try:
                    copying_and_modifying_slide.NewSlide(i, ppt, outputPres)
                except Exception as e:
                    print(f"Error adding activity slide {i}: {e}")

    # Last slides
    for i in range(8, 15):
        try:
            if form_data["english_version"] == True:
                ppt = print_link("Slides_DEBUT_FIN_ANGLAIS.pptx", "ANGLAIS")
            else:
                ppt = print_link("Slides_DEBUT_FIN_FRANCAIS.pptx", "FRANCAIS")
            copying_and_modifying_slide.NewSlide(i, ppt, outputPres)
        except Exception as e:
            print(f"Error adding slide {i}: {e}")
    
    # Save the PowerPoint file
    ppt_file = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("Fichiers PowerPoint", "*.pptx")])
    if ppt_file:
        try:
            outputPres.save(ppt_file)
            confirmation_label.config(text=f"Présentation enregistrée à {ppt_file}", fg="green")
        except Exception as e:
            print(f"Error saving presentation: {e}")




# Rest of the Tkinter interface code (unchanged from your original)
# Includes the day tabs creation, hotel dropdowns, and other functions
# ...



# Function to browse and select a logo file
def browse_logo():
    logo_path = filedialog.askopenfilename(filetypes=[("Fichiers d'images", "*.png;*.jpg;*.jpeg")])
    logo_entry.delete(0, tk.END)
    logo_entry.insert(0, logo_path)

# Function to add a new dropdown for hotel selection and a room number entry
def add_hotel_dropdown():
    global hotel_vars, room_number_vars

    # Create a new hotel dropdown
    new_hotel_var = tk.StringVar(window)
    new_hotel_var.set("Sélectionnez un hôtel")
    hotel_vars.append(new_hotel_var)  # Add the new hotel_var to the list

    # Place hotel dropdown in a new column
    col_position = len(hotel_vars) - 1  # Dynamic column placement
    new_hotel_menu = tk.OptionMenu(hotel_frame, new_hotel_var, *hotel_options)
    new_hotel_menu.grid(row=0, column=col_position, padx=5, pady=5)

    # Create an entry for the number of rooms corresponding to this hotel
    room_number_var = tk.StringVar(window)
    room_number_var.set("1")  # Default value
    room_number_vars.append(room_number_var)

    # Place the "Nombre de chambres" entry below the hotel dropdown
    room_label = tk.Label(hotel_frame, text="Nombre de chambres :")
    room_label.grid(row=1, column=col_position, padx=5, pady=5)
    room_entry = tk.Entry(hotel_frame, textvariable=room_number_var, width=15)
    room_entry.grid(row=2, column=col_position, padx=5, pady=5)

# Function to reset all hotel selections
def reset_hotels():
    global hotel_vars, room_number_vars

    # Clear the existing hotel and room number lists
    hotel_vars.clear()
    room_number_vars.clear()

    # Remove all hotel-related widgets from the hotel_frame
    for widget in hotel_frame.winfo_children():
        widget.destroy()

    # Reinitialize with the first hotel dropdown
    add_hotel_dropdown()

# Function to create tabs based on the number of days entered
def create_day_tabs():
    global day_selections, day_dates
    # Get the number of days
    num_days = days_entry.get()
    
    try:
        num_days = int(num_days)
    except ValueError:
        confirmation_label.config(text="Erreur : Veuillez entrer un nombre valide de jours.", fg="red")
        return
    
    if num_days <= 0:
        confirmation_label.config(text="Erreur : Le nombre de jours doit être supérieur à 0.", fg="red")
        return
    
    # Remove any existing day tabs
    for tab in day_frames:
        notebook.forget(tab)
    day_frames.clear()
    day_selections = []
    day_dates = []

    # Create new day tabs based on the number of days
    for day in range(1, num_days + 1):
        day_frame = tk.Frame(notebook)
        notebook.add(day_frame, text=f"Jour {day}")
        day_frames.append(day_frame)

        # Create an entry field for the date of the day (above the day's plan)
        tk.Label(day_frame, text=f"Date for Day {day}:").pack(pady=5)
        day_date_entry = tk.Entry(day_frame, width=50)
        day_date_entry.pack(pady=5)
        day_dates.append(day_date_entry)

        # Call the function from day_plan.py to organize the day plan for each tab
        day_plan.create_day_plan(day_frame, day - 1, day_selections)

        # Add an empty list to hold the steps selected for the day
        day_selections.append([])

    confirmation_label.config(text=f"{num_days} onglets créés pour les jours de voyage.", fg="green")


# Create the Tkinter window
window = tk.Tk()
window.title("PPT Marrakech Préférence-Événements")

# Create the notebook for day tabs
notebook = ttk.Notebook(window)
notebook.pack(pady=10, fill="both", expand=True)

# Create the main tab for the home page (Page principale)
main_frame = tk.Frame(notebook)
notebook.add(main_frame, text="Page principale")

# Add the checkbox for "Version en anglais" directly below the client entry
english_version_var = tk.BooleanVar()
english_version_checkbox = tk.Checkbutton(main_frame, text="Version en anglais", variable=english_version_var)
english_version_checkbox.pack(anchor='w', padx=10, pady=5)  # Align to the left

# Client input
tk.Label(main_frame, text="Client :").pack(pady=5)
client_entry = tk.Entry(main_frame, width=40)
client_entry.pack(pady=5)

# Type d'événement input
tk.Label(main_frame, text="Type d'événement :").pack(pady=5)
event_type_entry = tk.Entry(main_frame, width=40)
event_type_entry.pack(pady=5)

# Logo input
tk.Label(main_frame, text="Logo :").pack(pady=5)
logo_entry = tk.Entry(main_frame, width=40)
logo_entry.pack(pady=5)
browse_logo_button = tk.Button(main_frame, text="Parcourir Logo", command=browse_logo)
browse_logo_button.pack(pady=5)

# Hotel selection
tk.Label(main_frame, text="Choisir un hôtel :").pack(pady=5)
hotel_frame = tk.Frame(main_frame)
hotel_frame.pack(pady=5)

# Predefined hotel options
hotel_options = ["LES JARDINS DE LA KOUTOUBIA 5*", "SOFITEL MARRAKECH PALAIS IMPERIAL 5*", "BARCELO PALMERAIE 5*", "KENZI ROSE GARDEN HOTEL 5*"]

# Initialize the first hotel dropdown menu and room number entry
add_hotel_dropdown()

# Button to add a new hotel dropdown and room number entry
button_frame = tk.Frame(main_frame)
button_frame.pack(pady=5)

add_hotel_button = tk.Button(button_frame, text="Ajouter un hôtel +", command=add_hotel_dropdown)
add_hotel_button.pack(side=tk.LEFT, padx=5)

# Button to reset all hotels
reset_button = tk.Button(button_frame, text="Réinitialiser les hôtels", command=reset_hotels)
reset_button.pack(side=tk.LEFT, padx=5)

# Frame to align number of days, nights, people, and dates
days_nights_frame = tk.Frame(main_frame)
days_nights_frame.pack(pady=5)

# Number of Days input
tk.Label(days_nights_frame, text="Nombre de Jours :").grid(row=0, column=0, padx=5, pady=5)  # Label above the entry
days_entry = tk.Entry(days_nights_frame, width=15)
days_entry.grid(row=1, column=0, padx=5, pady=5)  # Entry below the label

# Number of Nights input
tk.Label(days_nights_frame, text="Nombre de Nuits :").grid(row=0, column=1, padx=5, pady=5)  # Label above the entry
nights_entry = tk.Entry(days_nights_frame, width=15)
nights_entry.grid(row=1, column=1, padx=5, pady=5)  # Entry below the label

# Number of People input
tk.Label(days_nights_frame, text="Nombre de Personnes :").grid(row=0, column=2, padx=5, pady=5)  # Label above the entry
people_entry = tk.Entry(days_nights_frame, width=15)
people_entry.grid(row=1, column=2, padx=5, pady=5)  # Entry below the label

# Dates input
tk.Label(days_nights_frame, text="Dates :").grid(row=0, column=3, padx=5, pady=5)  # Label above the entry
dates_entry = tk.Entry(days_nights_frame, width=15)  # Adjusted width to be the same size
dates_entry.grid(row=1, column=3, padx=5, pady=5)  # Entry below the label

# Button to validate and create tabs for each day
validate_button = tk.Button(main_frame, text="Valider et créer les jours", command=create_day_tabs)
validate_button.pack(pady=10)

# Generate PPT button
generate_button = tk.Button(main_frame, text="Générer PowerPoint", command=generate_ppt)
generate_button.pack(pady=10)

# Confirmation message
confirmation_label = tk.Label(main_frame, text="")
confirmation_label.pack(pady=10)


# Start the Tkinter loop
window.mainloop()
