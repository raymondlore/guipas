from datetime import datetime, timedelta

# Record the start time
start_time = datetime.now()

# Print the current date and time
print("Current Date and Time:", start_time.strftime("%Y-%m-%d %H:%M:%S"))

# Simulate some processing time
input("Press Enter to stop and calculate the running period...")

# Record the end time
end_time = datetime.now()

# Calculate the running period
running_period = end_time - start_time

# Print the running period
print("Running Period:", str(running_period))