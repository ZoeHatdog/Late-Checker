from datetime import datetime

def decimal_to_time(input_value):
    # Check if the input is a string and in the time format "%I:%M %p"
    if isinstance(input_value, str):
        time_format = "%I:%M %p"
        try:
            datetime.strptime(input_value, time_format)
            return input_value
        except ValueError:
            pass  # If the input string is not in time format, we will treat it as a decimal
    
    # If the input is not a valid time string, treat it as a decimal
    decimal = float(input_value)
    hours, remainder = divmod(decimal * 24, 1)
    minutes = remainder * 60
    return datetime.strptime(f"{int(hours)}:{int(minutes)}", "%H:%M").strftime("%I:%M %p")

# Example usage:
input_number = 0.5  # This should convert to "12:00 PM"
expectedtime = "8:00 AM"  # This should remain as "02:30 PM"
timein = "7:53 AM"
print(decimal_to_time(input_number))  # Output: "12:00 PM"
print(decimal_to_time(expectedtime))  # Output: "02:30 PM"

#### TIME TEST 


def calculate_minutes_late(actual_time_str, expected_time_str):
    try:
        actual_time = datetime.strptime(actual_time_str, "%I:%M %p")
        expected_time = datetime.strptime(expected_time_str, "%I:%M %p")
        difference = (actual_time - expected_time).total_seconds() / 60
        return max(0, difference)  # Ensure the difference is not negative
    except ValueError:
        return 0

if expectedtime >= timein: # MAS MAAGA MAS MALAKI VALUE , MAS MALATE MAS MALIIT VALUE
    print("Minutes Late: 0")
else:
    print(calculate_minutes_late(timein,expectedtime))









time = "7:41 AM"
time_in = datetime.strptime(time, "%I:%M %p")
formatted_time = time_in.strftime("%I:%M %p")
print(formatted_time)



##### TIME OUT 

print("TIMEOUT")
actual_timeout = "7:53 AM"
expected_timeout = "8:00 AM"

if expected_timeout < timein:
    print("TRUE")
else:
    print(calculate_minutes_late(actual_timeout,expected_timeout))
