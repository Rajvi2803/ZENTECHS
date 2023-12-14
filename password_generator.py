import random
import string

def important(length):
    letter = string.ascii_letters
    numb = string.digits
    symb = string.punctuation
    characs = letter + numb + symb  
    password = "".join(random.choice(characs) for _ in range(length))
    return password

def size():
    while True:
        try:
            length = int(input("Enter number of characters: "))
            if length <= 0:
                raise ValueError("Length should be greater than 0.")
            return length
        except ValueError:
            print("Invalid input. Please enter again.")


password_length = size()

try:
    generated_password = important(password_length)
    print(f"Generated Password: {generated_password}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
