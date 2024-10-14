import tkinter as tk

def create_day_plan(day_frame, day_index, day_selections):
    # List to store steps and labels for selections
    steps = []
    step_labels = []
    lines = []  # To store lines between steps

    # Initialize a frame for the steps (formerly cubes)
    step_frame = tk.Frame(day_frame)
    step_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)  # Center vertically and horizontally

    # Create a canvas for the fixed horizontal line (timeframe)
    line_canvas = tk.Canvas(day_frame, width=700, height=2)
    line_canvas.create_line(0, 1, 700, 1, fill="black")
    line_canvas.place(relx=0.5, rely=0.4, anchor=tk.CENTER)  # Adjust positioning

    # Frame for menus under the selected step
    menu_frame = tk.Frame(day_frame)
    menu_frame.place(relx=0.5, rely=0.6, anchor=tk.CENTER)  # Centered as well

    # Selected step index
    selected_step_index = None

    # Ensure there's an entry in day_selections for this day
    if len(day_selections) <= day_index:
        day_selections.append([])

    # Function to add a step
    def add_step():
        nonlocal selected_step_index
        step_index = len(steps)

        # Add a button as a step (above the line)
        step_button = tk.Button(step_frame, text=f"Étape {step_index + 1}", width=10, command=lambda idx=step_index: select_step(idx))
        step_button.grid(row=0, column=2 * step_index, padx=10)
        steps.append(step_button)

        # Add a label under the step to display the selected option
        selection_label = tk.Label(step_frame, text=" ")
        selection_label.grid(row=1, column=2 * step_index, padx=10)
        step_labels.append(selection_label)

        # Add a line between steps for aesthetics (except the first step)
        if step_index > 0:
            connector_line = tk.Canvas(step_frame, width=50, height=2)
            connector_line.create_line(0, 1, 50, 1, fill="black")
            connector_line.grid(row=0, column=2 * step_index - 1)
            lines.append(connector_line)

        selected_step_index = step_index
        update_menus()

    # Function to remove the last step
    def remove_step():
        if steps:
            steps[-1].destroy()
            step_labels[-1].destroy()
            steps.pop()
            step_labels.pop()
            clear_menus()

            if lines:
                lines[-1].destroy()
                lines.pop()

    # Function to select a step and update the menus
    def select_step(index):
        nonlocal selected_step_index
        selected_step_index = index
        update_menus()

    # Function to update the menus based on the selected step
    def update_menus():
        clear_menus()
        if selected_step_index is not None:
            # Create dropdown menus for each category (horizontally)
            create_dropdown("Divers", ["ARRIVEE",  "DEPART", "ACTIVITES A LA CARTE", "PETIT DEJEUNER A L'HOTEL"])
            create_dropdown("Déjeuner", ["DEJEUNER A L'HOTEL", "LODGE DU DESERT", "KASBAH DU TOUBKAL"])
            create_dropdown("Activité", ["APRES MIDI LIBRE", "REUNION A L'HOTEL KOUTOUBIA", "REUNION A L'HOTEL SOFITEL",  "DEPART EN 4X4", "DECOUVERTE DU DESERT EN 4X4", "COURS DE CUISINE MAROCAINE", 
                                         "TEMPS LIBRE DANS LES SOUKS", "CALECHE", "TREK DANS L'ATLAS"])
            create_dropdown("Soirée", ["DINER A L'HOTEL KOUTOUBIA", "DINER A L'HOTEL SOFITEL", "SOIREE DANS LE DESERT", "PALAIS JAD MAHAL", "DAR ZELLIJ", "PALAIS GHARNATA", "PALAIS DAR SOUKKAR", "AFTER AU BABOUCHKA"])

    # Function to clear all menus
    def clear_menus():
        for widget in menu_frame.winfo_children():
            widget.destroy()

    # Function to create a dropdown menu for a category
    def create_dropdown(label, options):
        frame = tk.Frame(menu_frame)
        frame.pack(side=tk.LEFT, padx=10)
        
        label_widget = tk.Label(frame, text=label)
        label_widget.pack()
        
        var = tk.StringVar(frame)
        var.set(options[0])  # default value
        dropdown = tk.OptionMenu(frame, var, *options, command=lambda value: update_selection(value))
        dropdown.pack()

    # Function to update the label under the selected step and store the selection in day_selections
    def update_selection(value):
        if selected_step_index is not None:
            step_labels[selected_step_index].config(text=value)
            # Store the selected step in day_selections
            if len(day_selections[day_index]) <= selected_step_index:
                day_selections[day_index].append(value)  # Add new step
            else:
                day_selections[day_index][selected_step_index] = value  # Update existing step

    # Buttons to add and remove steps (move them below the steps)
    button_frame = tk.Frame(day_frame)
    button_frame.place(relx=0.5, rely=0.45, anchor=tk.CENTER)

    add_button = tk.Button(button_frame, text="+", command=add_step)
    add_button.pack(side=tk.LEFT, padx=5)

    remove_button = tk.Button(button_frame, text="-", command=remove_step)
    remove_button.pack(side=tk.LEFT, padx=5)

    # Start with one step
    add_step()
