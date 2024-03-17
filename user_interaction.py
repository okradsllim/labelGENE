# user_interaction.py

import sys
import logging

def user_select_collection(collections):
    retry_count = 0
    max_retries = 10

    while retry_count < max_retries:
        try:
            print("\nPlease select a collection to process (or type 'q' to quit): \n")
            for i, collection in enumerate(collections, start=1):
                print(f"{i}. {collection['name']} - {collection['number']}")

            user_input = input("\nChoose number (or type 'q' to exit): \n\n")

            if user_input.lower() == 'q':
                print("\nExiting...Thanks for using the program, and have a great day!\n")
                sys.exit()

            selected_index = int(user_input) - 1
            logging.info(f"User selected index: {selected_index}")

            if 0 <= selected_index < len(collections):
                selected_collection = collections[selected_index]
                return selected_collection

            print("Invalid selection. Please enter a number from the list or 'q' to quit.")
            retry_count += 1

        except ValueError:
            print("Invalid input. Please enter a valid number or type 'q' to quit.\n")
            logging.warning("Invalid input encountered.\n")
            retry_count += 1

    logging.error("Maximum retries reached. Exiting program.")
    print("Too many incorrect attempts. Exiting program. Thanks for your using the program, and have a great day!\n")
    sys.exit()

def display_options(options_list, title):
    ''' Display options in order (generic for both series and box), with a custom display for lists over 20 items. '''
    try:
        logging.info(f"Displaying available {title} options:")
        
        # Determine the prefix and header based on the title
        if title.lower() == "box":
            prefix = "Box "
            header = "\nSelect box(es):"
        else:  # Assume it's series if not box
            prefix = ""
            header = "\nSelect series:"

        print(header)
        print()

        if len(options_list) > 30:
            # Display the first 10 options
            for i in range(10):
                print(f"{i + 1}. {prefix}{options_list[i]}")
            # Display ellipsis lines
            for _ in range(3):
                print("... ...")
            # Display the last 3 options
            for i in range(-3, 0, 1):
                box_number = options_list[i]
                print(f"{len(options_list) + i + 1}. {prefix}{box_number}")
                if box_number == '10001':
                    print("Note: Box 10001 is used as a flag. Please verify box number data before printing labels.")
        else:
            # If 30 or fewer, just display all options
            for i, option in enumerate(options_list, 1):
                print(f"{i}. {prefix}{option}")
                if option == '10001':
                    print("Note: Box 10001 is used as a flag. Please verify box number data before printing labels.")
                    
        print()  # Final newline for formatting
    except Exception as e:
        logging.error(f"Error displaying {title} options: {str(e)}")

def parse_user_input(input_str, options_list):
    ''' Parse user input as indices and return the corresponding values from options_list. '''
    selected_options_set = set()
    try:
        inputs = input_str.split(',')
        for input_item in inputs:
            input_item = input_item.strip()
            if '-' in input_item:
                start, end = map(int, input_item.split('-'))
                if 0 < start <= len(options_list) and 0 < end <= len(options_list):
                    selected_options_set.update(options_list[start - 1:end])
                else:
                    raise ValueError("Range selection out of bounds.")
            else:
                index = int(input_item)
                if 0 < index <= len(options_list):
                    selected_options_set.add(options_list[index - 1])
                else:
                    raise ValueError("Selection out of bounds.")
        return list(selected_options_set)
    except ValueError as e:
        logging.error(f"Error parsing user input: {str(e)}")
        print(f"Error: {str(e)}")
        return None