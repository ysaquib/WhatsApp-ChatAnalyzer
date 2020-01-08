# -*- coding: utf-8 -*-
import re
import xlsxwriter as xlw

f = open("texts.txt","r",encoding="utf-8")

SENDER = "Yusuf" # You
RECEIVER = "Sara" # Person you are chatting with

## DONT TOUCH THIS

wb = xlw.Workbook("chat_analysis.xlsx")
ws = wb.add_worksheet("Chat Data Analysis")

MSG_RCVD = 0
MSG_SENT = 0
IMG_RCVD = 0
IMG_SENT = 0
PREV_DAY = ""
CURR_DAY = ""
TOTAL_CHARS_SENT = 0
TOTAL_CHARS_RCVD = 0

TOTAL_WORDS_SENT = 0
TOTAL_WORDS_RCVD = 0

TOTAL_EMOJI_SENT = 0
TOTAL_EMOJI_RCVD = 0

CURR_MSNGR = ""

MSG_RCVD_TODAY = 0
MSG_SENT_TODAY = 0


def analyze(message, word_dict, emoji_dict, time_dict):

    # DECLARE GLOBALS -- DONT TOUCH
    global PREV_DAY
    global CURR_DAY
    global MSG_RCVD
    global MSG_RCVD_TODAY
    global MSG_SENT
    global MSG_SENT_TODAY
    global IMG_SENT
    global IMG_RCVD
    global TOTAL_CHARS_SENT
    global TOTAL_CHARS_RCVD
    global TOTAL_WORDS_SENT
    global TOTAL_WORDS_RCVD
    global TOTAL_EMOJI_SENT
    global TOTAL_EMOJI_RCVD
    global SENDER
    global RECEIVER
    global CURR_MSNGR


    ## IF THE DAY IS CURRENTLY ON GOING AND IT IS AT THE END OF THE FILE, THEN I NEED TO MAKE IT ANOTHER DAY!!!
    

    if re.match("\d+/\d+/\d+,", message):
        CURR_DAY = message.split(",")[0]
    if PREV_DAY == "":
        PREV_DAY = CURR_DAY
    if PREV_DAY != CURR_DAY:
        print ("    Messages Sent on " + PREV_DAY + ": " + str(MSG_SENT_TODAY))
        print ("Messages Received on " + PREV_DAY + ": " + str(MSG_RCVD_TODAY) + "\n")
        PREV_DAY = CURR_DAY
        MSG_SENT_TODAY = 0
        MSG_RCVD_TODAY = 0
    if " - " + SENDER + ": <Media omitted>" in message:
        IMG_SENT += 1
    if " - " + RECEIVER + ": <Media omitted>" in message:
        IMG_RCVD += 1

    str_msg = ""
    str_time = ""
    hour = ""
    str_emoji_regex = "[\U0001F100-\U0001F7EC]" #Encompasses All Emoji Unicode
    if " - " + SENDER + ":" in message:
        MSG_SENT += 1
        MSG_SENT_TODAY += 1
        CURR_MSNGR = SENDER
        TOTAL_CHARS_SENT += len(message.split(":")[1])
        str_msg = message.split(SENDER + ":")[1].strip()
        str_time = (message.split(",")[1].strip()).split("- " + SENDER)[0].strip()
        if str_msg == "<Media omitted>": str_msg = ""
        TOTAL_WORDS_SENT += len(re.findall("[^\w*'*\w]",  str_msg))
        TOTAL_EMOJI_SENT += len(re.findall(str_emoji_regex,  str_msg))
        time = "".join(re.findall("[\d\d:\d\d]", str_time))
        hour = time.split(":")[0]

    elif " - " + RECEIVER + ":" in message:
        MSG_RCVD += 1
        MSG_RCVD_TODAY += 1
        CURR_MSNGR = RECEIVER
        TOTAL_CHARS_RCVD += len(message.split(":")[1])
        str_msg = message.split(RECEIVER + ":")[1].strip()
        str_time = (message.split(",")[1].strip()).split("- " + RECEIVER)[0].strip()
        if str_msg == "<Media omitted>": str_msg = ""
        TOTAL_WORDS_RCVD += len(re.findall("[^\w*'*\w]",  str_msg))
        TOTAL_EMOJI_RCVD += len(re.findall(str_emoji_regex,  str_msg))
        time = "".join(re.findall("[\d\d:\d\d]", str_time))
        hour = time.split(":")[0]

    else:
        if CURR_MSNGR == SENDER:
            TOTAL_CHARS_SENT += len(message)
            str_msg = message
            str_time = ""
            if str_msg == "<Media omitted>": str_msg = ""
            TOTAL_WORDS_SENT += len(re.findall("[^\w*'*\w]",  str_msg))
            TOTAL_EMOJI_SENT += len(re.findall(str_emoji_regex,  str_msg))

        elif CURR_MSNGR == RECEIVER:
            TOTAL_CHARS_RCVD += len(message)
            str_msg = message
            str_time = ""
            if str_msg == "<Media omitted>": str_msg = ""
            TOTAL_WORDS_RCVD += len(re.findall("[^\w*'*\w]",  str_msg))
            TOTAL_EMOJI_RCVD += len(re.findall(str_emoji_regex,  str_msg))

    wordList = re.sub("[^\w*'*\w]", " ",  str_msg).split()
    emojiList = re.findall(str_emoji_regex, str_msg)

    inds = [i for i, x in enumerate(wordList) if (x == "ll" or x =="re" or x=="ve")]
    for j in inds:
        if j - 1 >= 0:
            wordList[j] = str(wordList[j-1]) + str(wordList[j])
            j -= 1

    for word in wordList:
        if word.lower().strip() not in word_dict:
            word_dict[word.lower().strip()] = 1
        else:
            word_dict[word.lower().strip()] += 1

    if hour not in time_dict:
        if CURR_MSNGR == SENDER: time_dict[hour] = [1, 0, 1]
        else: time_dict[hour] = [0, 1, 1]
    else:
        if CURR_MSNGR == SENDER:
            time_dict[hour][0] += 1
            time_dict[hour][2] += 1
        else: 
            time_dict[hour][1] += 1
            time_dict[hour][2] += 1

    for emoji in emojiList:
        if emoji not in emoji_dict:
            emoji_dict[emoji] = 1
        else:
            emoji_dict[emoji] += 1

def fileAnalyzer(file):
    file.readline()
    string = file.readline()
    if "Messages to this chat and calls are now secured with end-to-end encryption. Tap for more info." in string:
        string = file.readline()
    word_dict = {}
    time_dict = {}
    emoji_dict = {}
    while(string != ""):
        analyze(string, word_dict, emoji_dict, time_dict)
        string = file.readline()
    #print (word_dict)
    del time_dict[""]
    words_sorted = sorted(word_dict.items(), reverse = True, key = lambda x: x[1])
    emojis_sorted = sorted(emoji_dict.items(), reverse = True, key = lambda z: z[1])

    f.close()
    print ("    Total Messages Sent: " + str(MSG_SENT))
    print ("Total Messages Received: " + str(MSG_RCVD) + "\n")

    print ("         Total Messages: " + str(MSG_RCVD + MSG_SENT) + "\n")

    print ("      Total Images Sent: " + str(IMG_SENT))
    print ("  Total Images Received: " + str(IMG_RCVD) + "\n")

    print ("           Total Images: " + str(IMG_RCVD + IMG_SENT) + "\n")

    print ("       Total Chars Sent: " + str(TOTAL_CHARS_SENT))
    print ("   Total Chars Received: " + str(TOTAL_CHARS_RCVD) + "\n")

    print ("       Total Characters: " + str(TOTAL_CHARS_RCVD + TOTAL_CHARS_SENT) + "\n")

    print ("       Total Words Sent: " + str(TOTAL_WORDS_SENT))
    print ("   Total Words Received: " + str(TOTAL_WORDS_RCVD) + "\n")

    print ("            Total Words: " + str(TOTAL_WORDS_RCVD + TOTAL_WORDS_SENT) + "\n")

    print ("    Avg Sent Msg Length: " + str(TOTAL_CHARS_SENT//MSG_SENT) + " Chars")
    print ("Avg Received Msg Length: " + str(TOTAL_CHARS_RCVD//MSG_RCVD) + " Chars" + "\n")

    print ("100 Most Used Words: ")
    words100 = []
    for i in range(100):
        if not words_sorted[i][0].isdigit():
            print ("    " + str(i+1) + ". " + words_sorted[i][0] + " - " + str(words_sorted[i][1]))
            words100.append((words_sorted[i][0], words_sorted[i][1]))
    print ("\n")

    print ("      Total Emojis Sent: " + str(TOTAL_EMOJI_SENT))
    print (" Total Emojies Received: " + str(TOTAL_EMOJI_RCVD) + "\n")

    print ("          Total Emojies: " + str(TOTAL_EMOJI_SENT+TOTAL_EMOJI_RCVD) + "\n")
    
    print ("Most Used Emojis: ")
    emojis_all = []
    for j in range(len(emojis_sorted)):
        print ("    " + str(j+1)+". " + emojis_sorted[j][0] + " - " + str(emojis_sorted[j][1]))
        emojis_all.append((emojis_sorted[j][0], emojis_sorted[j][1]))
    print ("\n")
    list_times = ["00","01","02","03","04","05","06","07","08","09","10","11",
                    "12","13","14","15","16","17","18","19","20","21","22","23"]
    print ("  Message Freq per Hour: |   SENT   | RECEIVED |   TOTAL   ")
    msg_freq_hr = []
    for hr in list_times:
        print ("          "+hr+":00 - "+hr+":59 ::   " + str(time_dict[hr][0]) 
            + "   |   " + str(time_dict[hr][1]) 
            + "   |   " + str(time_dict[hr][2]))
        msg_freq_hr.append(
            (""+hr+":00 - "+hr+":59",
            int(time_dict[hr][0]),
            int(time_dict[hr][1]),
            int(time_dict[hr][2]))
        )
    

    analysis_list = [
        ("Total Messages Sent", int(MSG_SENT)),
        ("Total Messages Received", int(MSG_RCVD)),
        ("Total Messages", int(MSG_RCVD + MSG_SENT)),
        ("",""),
        ("Total Images Sent", int(IMG_SENT)),
        ("Total Images Received", int(IMG_RCVD)),
        ("Total Images", int(IMG_RCVD + IMG_SENT)),
        ("",""),
        ("Total Chars Sent", int(TOTAL_CHARS_SENT)),
        ("Total Chars Received", int(TOTAL_CHARS_RCVD)),
        ("Total Characters", int(TOTAL_CHARS_RCVD + TOTAL_CHARS_SENT)),
        ("",""),
        ("Total Words Sent", int(TOTAL_WORDS_SENT)),
        ("Total Words Received", int(TOTAL_WORDS_RCVD)),
        ("Total Words", int(TOTAL_WORDS_RCVD + TOTAL_WORDS_SENT)),
        ("",""),
        ("Avg Sent Msg Length", int(TOTAL_CHARS_SENT//MSG_SENT)),
        ("Avg Received Msg Length", int(TOTAL_CHARS_RCVD//MSG_RCVD)),
        ("",""),
        ("Total Emojis Sent", int(TOTAL_EMOJI_SENT)),
        ("Total Emojies Received", int(TOTAL_EMOJI_RCVD)),
        ("Total Emojies", int(TOTAL_EMOJI_SENT+TOTAL_EMOJI_RCVD))
    ]

    ## Write the data to the excel file
    row = 0
    col = 0
    for analysis, number in analysis_list:
        ws.write(row, col, analysis)
        ws.write(row, col+1, number)
        row += 1

    row = 0
    col = 3
    ws.write(row, col, "Most Used Words")
    ws.write(row, col+1, "Frequency")

    row = 1
    col = 3
    for word, freq in words100:
        ws.write(row, col, word)
        ws.write(row, col+1, freq)
        row += 1

    row = 0
    col = 6
    ws.write(row, col, "Most Used Emojis")
    ws.write(row, col+1, "Frequency")

    row = 1
    col = 6
    for emoji, freq in emojis_all:
        ws.write(row, col, emoji)
        ws.write(row, col+1, freq)
        row += 1

    row = 0
    col = 9
    ws.write(row, col, "Messages Per Hour")
    ws.write(row, col+1, "Sent")
    ws.write(row, col+2, "Received")
    ws.write(row, col+3, "Total")

    row = 1
    col = 9
    for hour, sent, recv, total in msg_freq_hr:
        ws.write(row, col, hour)
        ws.write(row, col+1, sent)
        ws.write(row, col+2, recv)
        ws.write(row, col+3, total)
        row += 1

    wb.close()

fileAnalyzer(f)