import win32com.client

# Create an instance of CANoe using its ProgID
canoe = win32com.client.Dispatch("CANoe.Application.2")

# You can now interact with CANoe. For example:
# Start CANoe
canoe.Start()

# Perform other operations, such as opening a configuration file
# canoe.Configuration.Load(r"C:\path\to\your\configuration.cfg")

# Stop CANoe
canoe.Stop()