# Native Python Libraries.
import psutil
import os
import csv
import time
from datetime import datetime

# Third Party Python Libraries.

# Self Built Python Libraries.


def performanceMonitor(directory: str, sleepTime: int = 5) -> None:
    """
    Monitors the system performance (RAM, CPU, and Disk usage) every default 5
    seconds and records the data into a CSV file. The CSV file is named based on
    the current date and time.

    Args:
        directory (str): The directory where the CSV file will be saved.

    Returns:
        None
    """
    # Get the current date and time for the CSV file name
    fileName = datetime.now().strftime("%Y-%m-%d-%H-%M-%S") + ".csv"
    filePath = os.path.join(directory, fileName)

    # Write the CSV headers
    with open(filePath, mode="w", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(
            [
                "processID",
                "dateTime",
                "hourExecution",
                "ramUse",
                "processorUse",
                "totalDiscUse",
            ]
        )

    try:
        # Get the current process ID
        pid = os.getpid()

        while True:

            # Get the current process information
            process = psutil.Process(pid)
            ramUse = process.memory_info().rss / (1024 * 1024)  # Convert to MB
            processorUse = psutil.cpu_percent(interval=1)
            totalDiscUse = psutil.disk_usage("/").percent  # For the root directory

            # Get the current date and time
            currentDateTime = datetime.now().strftime("%Y-%m-%d")
            currentHour = datetime.now().strftime("%H:%M:%S")

            # Append the performance data to the CSV file
            with open(filePath, mode="a", newline="") as file:
                writer = csv.writer(file)
                writer.writerow(
                    [
                        pid,
                        currentDateTime,
                        currentHour,
                        ramUse,
                        processorUse,
                        totalDiscUse,
                    ]
                )

            # Sleep for default 5 seconds before the next check
            time.sleep(sleepTime)

    except KeyboardInterrupt:
        # Stops monitoring when the user interrupts the program
        print(f"Performance monitoring stopped. Data saved to {filePath}")
