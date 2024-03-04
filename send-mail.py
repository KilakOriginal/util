from urllib.parse import urlparse

def read_configuration() -> dict:
    while True:
        smtp_address = input("Please enter the SMTP server address: ")
        if smtp_address:
            break
        print("SMTP server cannot be empty; try again.")
    domain = '.'.join(urlparse(smtp_address).netloc.split('.')[-2:])

    while True:
        user = f"{input('Please enter your user name: ')}@{domain}"
        if user:
            break
        print("User cannot be empty; try again.")
    
    while True:
        password = input("Please enter your password: ")
        if password:
            break
        print("Password cannot be empty; try again.")
    
    while True:
        try:
            port = int(input("[Optional] Please enter a port (Default: 465): "))
            if 1 <= port or 65535 >= port:
                break
            print("Invalid port number; try again.")
        except ValueError:
            print(f"'{port}' is not a number; try again.")

    return {"smtp_address" : smtp_address,
            "port"         : port,
            "user"         : user,
            "password"     : password}

def generate_email(template: str, replace_with: dict) -> list[str]:
    result = []
    keys = list(replace_with.keys())
    values = list(replace_with.values())
    size = len(values[0])
    for value in values:
        if len(value) != size:
            print(f"Warning: Not enough values in {value}")
            if len(value) < size:
                size = len(value)
    
    for i in range(size):
        text = template
        for j, key in enumerate(keys):
            text = text.replace(key, values[j][i])
        result.append(text)

    return result

def send_mail(to: str, subject: str, message: str, configuration: dict):
    import smtplib, ssl
    from socket import gaierror

    context = ssl.create_default_context()

    # Try sending an email
    try:
        with smtplib.SMTP_SSL(configuration["smtp_address"], configuration["port"], context=context) as server:
            server.login(configuration["user"], configuration["password"])
            server.sendmail(configuration["user"], to , f'Subject: [Kummerkasten] {subject}\n\n{message}'.encode("utf-8"))
        return True  # OK
    except (gaierror, ConnectionRefusedError):  # Unable to connect to the server specified in the configuration file
        print('Failed to connect to the server. Bad connection settings?')
    except smtplib.SMTPServerDisconnected:  # Unable to maintain connection; usually authentication error (username/password)
        print('Failed to connect to the server. Wrong user/password?')
    except smtplib.SMTPException as e:  # Unable to send email for some other reason
        print('SMTP error occurred: ' + str(e))
    return False  # Error

def main():
    text = "Hallo <name>! Deine Rolle ist <role>."
    replace_with = {"<name>" : ["Frank", "Sabine", "Klara", "Paul"],
                    "<role>" : ["Springer", "Bar", "Garderobe", "Türsteher"]}
    
    print(generate_email(text, replace_with))

if __name__ == "__main__":
    main()
