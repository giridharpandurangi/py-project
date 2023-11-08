import random
jackpot=random.randint(1,100)

number = int(input("guess the number between 1 to 100 :"))
while number != jackpot:
    if number < jackpot:
        print("Too low")
    else:
        print("Too high")
    number = int(input("guess the number between 1 to 100 :"))
else:
    print("You guessed it right")
    print("The number was :",jackpot)