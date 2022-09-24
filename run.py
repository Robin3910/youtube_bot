import json
import random
import time

from src.instadm import InstaDM

f = open('infos/accounts.json', )
accounts = json.load(f)

with open('infos/usernames.txt', 'r') as f:
    usernames = [line.strip() for line in f]

message = ["Beiens(a 10-year brand aiming at learning toys) is looking for brand partners for New Product Release! Beiens®mental math game is expected to be released in your country in November. The sales revenue of it in China has exceeded 10 million US dollars.　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣ 　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣　⁣⁣⁣⁣ 　⁣⁣⁣⁣　⁣⁣⁣⁣⁣⁣⁣ I reviewed your profile and believe it matches our criteria. As the brand partner you will: ✔️First to get FREE beiens®mental math game ✔️Receive promo code and earn commission on every sale　⁣⁣⁣✔️⁣⁣⁣⁣Get more traffic, fans, and commissions as the first launch influencers　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣✔️Be promoted on our social medias　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　⁣⁣⁣　⁣⁣⁣✔️Get a chance to become our brand ambassador",
           "Features of beiens®mental math game:　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　⁣⁣⁣　✔️Discover children's interest and improve ability of mental math　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　⁣⁣⁣　　⁣⁣⁣　✔️4 games, learning math through play　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　⁣⁣⁣✔️3 levels, 17 types of questions　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　✔️10 million+ question bank for pre-K to grade 6 (4-12 years)　✔️Automatic question generation and correctionLearn more: 　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　✔️ https://bit.ly/3AF0tA4　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　　⁣⁣⁣　⁣⁣⁣　⁣⁣⁣⁣⁣⁣　Please DM if interested! Or email to beiensmarketing@gmail.com"]


while True:
    if not usernames:
        print('Finished usernames.')
        break

    for account in accounts:
        if not usernames:
            break
        # Auto login
        insta = InstaDM(username=account["username"],
                        password=account["password"], headless=False)

        # while True:
        #     if not usernames:
        #         break
        #
        #     username = usernames.pop()
        #     # Send message
        #     insta.sendMessage(
        #         usernameser=username, message=message)
