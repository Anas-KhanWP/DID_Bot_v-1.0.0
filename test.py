import subprocess

def ping_server(server):
    try:
        # Start the ping process
        process = subprocess.Popen(['ping', server, '-t'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        # Continuously read the output of the ping command
        while True:
            output = process.stdout.readline()
            if output == b'' and process.poll() is not None:
                break
            if output:
                print(output.strip().decode())
                
    except KeyboardInterrupt:
        print("\nPing process interrupted by user.")
        process.terminate()
    except Exception as e:
        print(f"An error occurred: {e}")

# Ping the server
ping_server('mail.khanlocalseo.com')
