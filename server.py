import socket

# Set up TCP server
server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
server_address = ('localhost', 12345)
server_socket.bind(server_address)
server_socket.listen(1)

print('Server listening on {}:{}'.format(*server_address))

# Accept incoming connection
client_socket, client_address = server_socket.accept()
print('Connected to:', client_address)

# Receive data from Q-SYS Lua script
data = client_socket.recv(1024)
print('Received from Q-SYS:', data.decode())

# Process the PowerPoint or perform other tasks
# Replace this with your actual logic

# Prepare response
response = 'Speaker notes: This is a test response'

# Send response back to Q-SYS Lua script
client_socket.send(response.encode())

# Close the connection
client_socket.close()
server_socket.close()