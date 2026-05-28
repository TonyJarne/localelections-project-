import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart
import os

# ── RAW DATA ────────────────────────────────────────────────────────────────

WARDS = [
    # (ward, candidate, party, votes, elected)
    # Bearwood
    ("Bearwood", "Hilary Jane Grandey",        "Green Party",  1599, True),
    ("Bearwood", "John Ashley Tipper",          "Green Party",  1569, True),
    ("Bearwood", "Jennifer Lynn Hemingway",     "Labour Party", 1535, True),
    ("Bearwood", "Paul William Connor",         "Green Party",  1493, False),
    ("Bearwood", "Bobbie Fenton",               "Labour Party", 1472, False),
    ("Bearwood", "Sita Ram Malhi",              "Labour Party", 1306, False),
    ("Bearwood", "Paul Bithell",                "Reform UK",     818, False),
    ("Bearwood", "Nathan Luke Poole",           "Reform UK",     721, False),
    ("Bearwood", "Robert William White",        "Reform UK",     683, False),
    ("Bearwood", "Leon Barnfield",              "Conservative",  343, False),
    ("Bearwood", "Balbinder Kaur",              "Conservative",  261, False),
    ("Bearwood", "Karol Krupa",                 "Conservative",  201, False),
    ("Bearwood", "Richard Gingell",             "Other",          58, False),

    # Blackheath
    ("Blackheath", "Michael Andrew Cooper",     "Reform UK",    1985, True),
    ("Blackheath", "Dave Williams",             "Reform UK",    1800, True),
    ("Blackheath", "Mona Khurana",              "Reform UK",    1668, True),
    ("Blackheath", "Kerrie Carmichael",         "Labour Party", 1099, False),
    ("Blackheath", "Jag Singh",                 "Labour Party",  892, False),
    ("Blackheath", "Shobha Sharma",             "Labour Party",  860, False),
    ("Blackheath", "Deborah White",             "Green Party",   541, False),
    ("Blackheath", "Ellis Wood",                "Green Party",   477, False),
    ("Blackheath", "Richard Andrew Watson",     "Green Party",   443, False),
    ("Blackheath", "Gemma Barnfield",           "Conservative",  488, False),
    ("Blackheath", "Bhapinder Singh Bains",     "Conservative",  385, False),
    ("Blackheath", "Karolina Kordala",          "Conservative",  329, False),
    ("Blackheath", "Randall Shergill",          "Independent",   100, False),

    # Bristnall
    ("Bristnall", "Jonathan James Fox",         "Reform UK",    1283, True),
    ("Bristnall", "Amolak Singh Dhariwal",      "Labour Party", 1215, True),
    ("Bristnall", "Liane Winsper",              "Reform UK",    1148, True),
    ("Bristnall", "Ellen Fenton",               "Labour Party", 1138, False),
    ("Bristnall", "Ravinder Singh",             "Reform UK",    1131, False),
    ("Bristnall", "Tom Johnston",               "Labour Party", 1100, False),
    ("Bristnall", "Kalman Dean-Richards",       "Green Party",   676, False),
    ("Bristnall", "Louise Brunt",               "Green Party",   639, False),
    ("Bristnall", "Richard Touzjian",           "Green Party",   509, False),
    ("Bristnall", "Colin Blewitt",              "Conservative",  486, False),
    ("Bristnall", "Amran Singh Jaypal",         "Conservative",  392, False),
    ("Bristnall", "Rafal Dobiczek",             "Conservative",  343, False),

    # Charlemont & Grove Vale
    ("Charlemont & Grove Vale", "Bob Jones",            "Reform UK",    1499, True),
    ("Charlemont & Grove Vale", "Ray Nock",             "Reform UK",    1487, True),
    ("Charlemont & Grove Vale", "Rachael Mitchell",     "Reform UK",    1443, True),
    ("Charlemont & Grove Vale", "Dalel Singh Bhamra",   "Labour Party", 1131, False),
    ("Charlemont & Grove Vale", "David Dean Fisher",    "Conservative",  868, False),
    ("Charlemont & Grove Vale", "Satish Danny Paul",    "Labour Party",  864, False),
    ("Charlemont & Grove Vale", "Stephanie Thomas",     "Labour Party",  817, False),
    ("Charlemont & Grove Vale", "Aran Duggal",          "Conservative",  772, False),
    ("Charlemont & Grove Vale", "Amrita Jasmine Dunn",  "Conservative",  673, False),
    ("Charlemont & Grove Vale", "Parminder Kaur Bahia", "Green Party",   531, False),
    ("Charlemont & Grove Vale", "Maddi Lane",           "Green Party",   438, False),
    ("Charlemont & Grove Vale", "Primrose Joy Taylor",  "Green Party",   403, False),
    ("Charlemont & Grove Vale", "Bertram Richards",     "Lib Dem",       182, False),

    # Cradley Heath & Old Hill
    ("Cradley Heath & Old Hill", "Mark Johnson",        "Reform UK",    1482, True),
    ("Cradley Heath & Old Hill", "Craig Morris",        "Reform UK",    1430, True),
    ("Cradley Heath & Old Hill", "Nathan John Williams","Reform UK",    1392, True),
    ("Cradley Heath & Old Hill", "Julie Webb",          "Labour Party",  919, False),
    ("Cradley Heath & Old Hill", "Tom Knowles",         "Labour Party",  908, False),
    ("Cradley Heath & Old Hill", "Sheraz Khan",         "Labour Party",  836, False),
    ("Cradley Heath & Old Hill", "Katherine Anderson",  "Green Party",   566, False),
    ("Cradley Heath & Old Hill", "Brad Dymond",         "Green Party",   435, False),
    ("Cradley Heath & Old Hill", "Awais Qaisar",        "Green Party",   387, False),
    ("Cradley Heath & Old Hill", "Annetta Powell",      "Conservative",  348, False),
    ("Cradley Heath & Old Hill", "Kimberly Goddard",    "Conservative",  329, False),
    ("Cradley Heath & Old Hill", "Satinder Dunn",       "Conservative",  336, False),
    ("Cradley Heath & Old Hill", "Sam Richardson",      "Lib Dem",       230, False),

    # Friar Park & Stone Cross
    ("Friar Park & Stone Cross", "Nick Fawcett",        "Reform UK",    1671, True),
    ("Friar Park & Stone Cross", "Lisa Jane Weaver",    "Reform UK",    1598, True),
    ("Friar Park & Stone Cross", "Jack Sabharwal",      "Reform UK",    1574, True),
    ("Friar Park & Stone Cross", "Simon Hackett",       "Labour Party",  829, False),
    ("Friar Park & Stone Cross", "Terry Fitzgerald",    "Labour Party",  688, False),
    ("Friar Park & Stone Cross", "Elizabeth Ann Giles", "Labour Party",  674, False),
    ("Friar Park & Stone Cross", "Paul Barnfield",      "Conservative",  439, False),
    ("Friar Park & Stone Cross", "Dennis Oba",          "Conservative",  312, False),
    ("Friar Park & Stone Cross", "Agnieszka Wojtowicz", "Conservative",  274, False),
    ("Friar Park & Stone Cross", "Amy Pittaway",        "Green Party",   303, False),
    ("Friar Park & Stone Cross", "Cameron Robinson-Perry","Green Party", 253, False),
    ("Friar Park & Stone Cross", "Nita Leyshon",        "Green Party",   280, False),
    ("Friar Park & Stone Cross", "Colin Nicholas Rankine","Other",        38, False),

    # Great Barr, Tamebridge & Yew Tree
    ("Great Barr, Tamebridge & Yew Tree", "Connor Lee Horton",  "Labour Party", 1489, True),
    ("Great Barr, Tamebridge & Yew Tree", "Michaela Allcock",   "Labour Party", 1353, True),
    ("Great Barr, Tamebridge & Yew Tree", "Margaret Laura Sutton","Reform UK",  1276, True),
    ("Great Barr, Tamebridge & Yew Tree", "Ray Darby",          "Reform UK",   1258, False),
    ("Great Barr, Tamebridge & Yew Tree", "Darren Harding",     "Reform UK",   1219, False),
    ("Great Barr, Tamebridge & Yew Tree", "Mazhar Hussain",     "Labour Party", 1108, False),
    ("Great Barr, Tamebridge & Yew Tree", "Adrian Frederick Jones","Conservative", 598, False),
    ("Great Barr, Tamebridge & Yew Tree", "Vik Chandla",        "Green Party",  503, False),
    ("Great Barr, Tamebridge & Yew Tree", "Debi Hayward",       "Green Party",  458, False),
    ("Great Barr, Tamebridge & Yew Tree", "Jaroslaw Cwik",      "Conservative", 471, False),
    ("Great Barr, Tamebridge & Yew Tree", "Frank Emukperuo",    "Green Party",  402, False),
    ("Great Barr, Tamebridge & Yew Tree", "Michal Lorek",       "Conservative", 418, False),
    ("Great Barr, Tamebridge & Yew Tree", "Mark Smith",         "Lib Dem",      335, False),
    ("Great Barr, Tamebridge & Yew Tree", "Akramul Hoque",      "Lib Dem",      282, False),

    # Great Bridge
    ("Great Bridge", "Keith Stephen Edge",      "Reform UK",    1591, True),
    ("Great Bridge", "Mark Terence Webb",        "Reform UK",    1409, True),
    ("Great Bridge", "Brad Steven Simms",        "Reform UK",    1380, True),
    ("Great Bridge", "Kartar Singh Dosanjh",     "Labour Party",  843, False),
    ("Great Bridge", "Will Gill",                "Green Party",   836, False),
    ("Great Bridge", "Sahdaish Kaur Pall",       "Labour Party",  724, False),
    ("Great Bridge", "Soyfur Rahman",            "Labour Party",  680, False),
    ("Great Bridge", "Joe Cogavin",              "Green Party",   621, False),
    ("Great Bridge", "Naseer Hussain",           "Green Party",   610, False),
    ("Great Bridge", "Aaron David Emms",         "Conservative",  283, False),
    ("Great Bridge", "Parveen Ahkter",           "Conservative",  273, False),
    ("Great Bridge", "Shahid Mahmood",           "Conservative",  183, False),

    # Greets Green & Lyng
    ("Greets Green & Lyng", "Mohammed Ahad",     "Labour Party", 1597, True),
    ("Greets Green & Lyng", "Pam Randhawa",      "Labour Party", 1359, True),
    ("Greets Green & Lyng", "Jackie Taylor",     "Labour Party", 1356, True),
    ("Greets Green & Lyng", "Paul Green",        "Reform UK",     805, False),
    ("Greets Green & Lyng", "Mark McDermott",    "Reform UK",     773, False),
    ("Greets Green & Lyng", "Mike Stanyer",      "Reform UK",     760, False),
    ("Greets Green & Lyng", "Laura Curtis",      "Green Party",   484, False),
    ("Greets Green & Lyng", "Mohammed Hussain",  "Green Party",   480, False),
    ("Greets Green & Lyng", "Renee Collins",     "Green Party",   460, False),
    ("Greets Green & Lyng", "Patrycja Suchenek", "Conservative",  245, False),
    ("Greets Green & Lyng", "Dagmara Wojtowicz", "Conservative",  229, False),
    ("Greets Green & Lyng", "Patryk Wojtowicz",  "Conservative",  203, False),
    ("Greets Green & Lyng", "Balbir Singh Negi", "Lib Dem",       215, False),

    # Hateley Heath
    ("Hateley Heath", "Paul Moore",              "Labour Party", 1266, True),
    ("Hateley Heath", "Amardeep Singh",          "Labour Party", 1166, True),
    ("Hateley Heath", "Stephen James Fellows",   "Reform UK",    1127, True),
    ("Hateley Heath", "Dave Moore",              "Reform UK",    1088, False),
    ("Hateley Heath", "Doug Perry",              "Reform UK",    1062, False),
    ("Hateley Heath", "Jayne Wilkinson",         "Labour Party", 1073, False),
    ("Hateley Heath", "Theresa Millard",         "Green Party",   336, False),
    ("Hateley Heath", "Sunitha Soni",            "Green Party",   282, False),
    ("Hateley Heath", "Liegha Taylor",           "Green Party",   255, False),
    ("Hateley Heath", "Dorota Cwik",             "Conservative",  268, False),
    ("Hateley Heath", "Susan Gollins",           "Conservative",  266, False),
    ("Hateley Heath", "Ram Sarup",               "Conservative",  211, False),
    ("Hateley Heath", "Martin Roebuck",          "Lib Dem",        80, False),

    # Hill Top
    ("Hill Top", "Dean Adam Hollowood",          "Reform UK",    1381, True),
    ("Hill Top", "Amaan Husen",                  "Reform UK",    1193, True),
    ("Hill Top", "Aaron Khuttan",                "Reform UK",    1183, True),
    ("Hill Top", "Jenny Chidley",                "Labour Party", 1067, False),
    ("Hill Top", "Kulwant Singh Uppal",          "Labour Party", 1013, False),
    ("Hill Top", "Skerntian Keri",               "Labour Party",  884, False),
    ("Hill Top", "Steve Simcox",                 "Conservative",  391, False),
    ("Hill Top", "Zach Bates",                   "Green Party",   385, False),
    ("Hill Top", "Rohan Lal",                    "Conservative",  365, False),
    ("Hill Top", "Olivia Emily Shaw",            "Green Party",   345, False),
    ("Hill Top", "Al Imon Mohammed",             "Green Party",   270, False),
    ("Hill Top", "Oskar Wojtowicz",              "Conservative",  275, False),
    ("Hill Top", "Manjit Singh Lall",            "Lib Dem",       191, False),

    # Langley
    ("Langley", "Pete Durnell",                  "Reform UK",    1450, True),
    ("Langley", "Paul Christopher Leavey",       "Reform UK",    1382, True),
    ("Langley", "Tuli Zefi",                     "Reform UK",    1246, True),
    ("Langley", "Caroline Louise Owen",          "Labour Party",  992, False),
    ("Langley", "Jill Tromans",                  "Labour Party",  948, False),
    ("Langley", "Nadia Iqbal",                   "Labour Party",  935, False),
    ("Langley", "Eve Batten",                    "Green Party",   614, False),
    ("Langley", "James Deans",                   "Green Party",   550, False),
    ("Langley", "Sarah Louise Deans",            "Green Party",   533, False),
    ("Langley", "Sandra Collinge",               "Conservative",  463, False),
    ("Langley", "Izabela Blaszczak",             "Conservative",  367, False),
    ("Langley", "Krzysztof Wojtowicz",           "Conservative",  300, False),
    ("Langley", "Dheeraj Singh Rawat",           "Lib Dem",       205, False),

    # Newton & Valley
    ("Newton & Valley", "Keith Robert Allcock",  "Labour Party", 1311, True),
    ("Newton & Valley", "Kenny Jinks",           "Reform UK",    1073, True),
    ("Newton & Valley", "Elaine Mary Giles",     "Labour Party", 1068, True),
    ("Newton & Valley", "Rob Williams",          "Reform UK",    1053, False),
    ("Newton & Valley", "Tarjinder Singh Bassi", "Reform UK",    1029, False),
    ("Newton & Valley", "Saj Ashraf",            "Labour Party", 1021, False),
    ("Newton & Valley", "Liam Emanuel",          "Green Party",   548, False),
    ("Newton & Valley", "Jeremy David Parker",   "Green Party",   533, False),
    ("Newton & Valley", "Tiffany Sims",          "Green Party",   519, False),
    ("Newton & Valley", "Roman Lal",             "Conservative",  421, False),
    ("Newton & Valley", "Ewa Lorek",             "Conservative",  371, False),
    ("Newton & Valley", "Alina Mazur",           "Conservative",  318, False),
    ("Newton & Valley", "Amanda Jenkins",        "Lib Dem",       300, False),
    ("Newton & Valley", "Daljit Kaur",           "Lib Dem",       256, False),

    # Old Warley
    ("Old Warley", "Karl Terry Leech",           "Reform UK",    1379, True),
    ("Old Warley", "Connor Marshall",            "Reform UK",    1376, True),
    ("Old Warley", "Luke Dean Cotterill",        "Labour Party", 1288, True),
    ("Old Warley", "Harnoor Bhullar",            "Labour Party", 1206, False),
    ("Old Warley", "Baljinder Singh",            "Reform UK",    1201, False),
    ("Old Warley", "Chippie Kalebe-Nyamongo",    "Labour Party", 1036, False),
    ("Old Warley", "Mark Steven Holdroyd",       "Green Party",   596, False),
    ("Old Warley", "Tanisha Reid",               "Green Party",   559, False),
    ("Old Warley", "Mohammed Mahmood Amin Miah", "Green Party",   468, False),
    ("Old Warley", "Eric Dawes",                 "Conservative",  501, False),
    ("Old Warley", "Aleksandra Rokita",          "Conservative",  324, False),
    ("Old Warley", "Katarzyna Sikora",           "Conservative",  320, False),
    ("Old Warley", "Bob Smith",                  "Lib Dem",       258, False),

    # Oldbury
    ("Oldbury", "Nagi Daya Singh",               "Labour Party", 1177, True),
    ("Oldbury", "Suzanne Maria Hartwell",        "Labour Party", 1105, True),
    ("Oldbury", "Rizwan Jalil",                  "Labour Party", 1089, True),
    ("Oldbury", "Stuart Hill",                   "Reform UK",     767, False),
    ("Oldbury", "Rita Randell",                  "Reform UK",     745, False),
    ("Oldbury", "Andrew John Such",              "Reform UK",     730, False),
    ("Oldbury", "Jake Simon Cree",               "Green Party",   486, False),
    ("Oldbury", "Andy Dangerfield",              "Green Party",   452, False),
    ("Oldbury", "Thomas Alan Gerrish",           "Green Party",   403, False),
    ("Oldbury", "Amit Amit",                     "Conservative",  387, False),
    ("Oldbury", "Daniel Marquez",                "Conservative",  306, False),
    ("Oldbury", "Shazia Khan",                   "Conservative",  262, False),

    # Princes End
    ("Princes End", "Gary Andrew Dale",          "Reform UK",    1569, True),
    ("Princes End", "Neil Shirvington",          "Reform UK",    1472, True),
    ("Princes End", "Geoffrey Lionel Sutton",    "Reform UK",    1394, True),
    ("Princes End", "Archer Williams",           "Labour Party",  687, False),
    ("Princes End", "Geoff Deakin",              "Labour Party",  562, False),
    ("Princes End", "Joanna Quaye",              "Labour Party",  450, False),
    ("Princes End", "Kelly Cranston",            "Conservative",  486, False),
    ("Princes End", "Justyna Kordala",           "Conservative",  345, False),
    ("Princes End", "Natalie Weston",            "Conservative",  280, False),
    ("Princes End", "Helen Broome",              "Green Party",   298, False),
    ("Princes End", "Seona Deuchar",             "Green Party",   192, False),
    ("Princes End", "Shareen Khan",              "Green Party",   197, False),
    ("Princes End", "Chandra Shekhar Sharma",    "Lib Dem",       126, False),
    ("Princes End", "David John Wilkes",         "Other",         113, False),

    # Rowley
    ("Rowley", "Paul Tromans",                   "Reform UK",    1677, True),
    ("Rowley", "Ritchie Colin Massey",           "Reform UK",    1616, True),
    ("Rowley", "Jeet Taheem",                    "Reform UK",    1402, True),
    ("Rowley", "Sohail Iqbal",                   "Labour Party",  945, False),
    ("Rowley", "Khayam Khan",                    "Labour Party",  871, False),
    ("Rowley", "Desta Harris",                   "Labour Party",  852, False),
    ("Rowley", "Kelly Burton",                   "Conservative",  458, False),
    ("Rowley", "Benjamin James Morris",          "Green Party",   350, False),
    ("Rowley", "Joseph Loudon",                  "Green Party",   339, False),
    ("Rowley", "Eva Echo",                       "Green Party",   401, False),
    ("Rowley", "Sharnjit Kaur",                  "Conservative",  326, False),
    ("Rowley", "Ryan Lal",                       "Conservative",  284, False),

    # Smethwick
    ("Smethwick", "Parbinder Kaur",              "Labour Party", 1339, True),
    ("Smethwick", "Luke John Davies",            "Labour Party", 1298, True),
    ("Smethwick", "Ash Lewis",                   "Labour Party", 1182, True),
    ("Smethwick", "Kieran David John Diver",     "Green Party",   673, False),
    ("Smethwick", "Mark Redding",                "Green Party",   642, False),
    ("Smethwick", "Kez Sleeman",                 "Green Party",   593, False),
    ("Smethwick", "Simran Kaur",                 "Reform UK",     605, False),
    ("Smethwick", "Mamta Bhardwaj",              "Reform UK",     603, False),
    ("Smethwick", "Dwayne Robert James McQuaid", "Reform UK",     570, False),
    ("Smethwick", "Harjindar Kaur Marwaha",      "Conservative",  239, False),
    ("Smethwick", "Tarlochan Sunner",            "Conservative",  185, False),
    ("Smethwick", "Ann Wylie",                   "Conservative",  192, False),

    # Soho & Victoria
    ("Soho & Victoria", "Ragih Saleh Ahmed Muflihi", "Labour Party", 1304, True),
    ("Soho & Victoria", "Farut Shaeen",          "Labour Party", 1275, True),
    ("Soho & Victoria", "Mohammed Jalal Uddin",  "Labour Party", 1232, True),
    ("Soho & Victoria", "Andrea Melissa Boxall", "Green Party",   676, False),
    ("Soho & Victoria", "Erin Lewis",            "Green Party",   613, False),
    ("Soho & Victoria", "Stefan Joseph Gareth Smith","Green Party",581, False),
    ("Soho & Victoria", "Christopher Clemson",   "Reform UK",     260, False),
    ("Soho & Victoria", "David Michael Jones",   "Reform UK",     254, False),
    ("Soho & Victoria", "Les Smith",             "Reform UK",     220, False),
    ("Soho & Victoria", "Susan Neale",           "Conservative",  258, False),
    ("Soho & Victoria", "Jolanta Obcowska-Sarup","Conservative",  203, False),
    ("Soho & Victoria", "Leslie Trumpeter",      "Conservative",  200, False),

    # St Paul's
    ("St Paul's", "Aqeela Choudhry",             "Labour Party", 1745, True),
    ("St Paul's", "Abid Hussain",                "Labour Party", 1650, True),
    ("St Paul's", "Charn Singh Padda",           "Labour Party", 1620, True),
    ("St Paul's", "Owais Sajed Khan",            "Green Party",   646, False),
    ("St Paul's", "Kevin Priest",                "Green Party",   517, False),
    ("St Paul's", "Kenan Taylor",                "Green Party",   462, False),
    ("St Paul's", "Sukhbir Singh Gill",          "Independent",   474, False),
    ("St Paul's", "Steven Finch",                "Reform UK",     301, False),
    ("St Paul's", "Mark John Pilkington",        "Reform UK",     242, False),
    ("St Paul's", "Lynne Tomkinson",             "Reform UK",     226, False),
    ("St Paul's", "Shirley Andrea Johnson",      "Conservative",  224, False),
    ("St Paul's", "Jay Singh",                   "Other",         214, False),
    ("St Paul's", "Anna Misiewicz",              "Conservative",  180, False),
    ("St Paul's", "Krzysztof Misiewicz",         "Conservative",  135, False),

    # Tipton Green
    ("Tipton Green", "Richard James Jeffcoat",   "Independent",  1443, True),
    ("Tipton Green", "Tim Hordley",              "Reform UK",    1386, True),
    ("Tipton Green", "Matt Lloyd",               "Reform UK",    1345, True),
    ("Tipton Green", "Syeda Khatun",             "Labour Party", 1101, False),
    ("Tipton Green", "Khurshid Haque",           "Labour Party", 1007, False),
    ("Tipton Green", "Nahid Iqbal",              "Labour Party",  893, False),
    ("Tipton Green", "Raven Dixon-Biggs",        "Green Party",   554, False),
    ("Tipton Green", "Prabhkin Kaur Bhullar",    "Green Party",   535, False),
    ("Tipton Green", "Aldo Mussi",               "Green Party",   355, False),
    ("Tipton Green", "Tom Lewandowski",          "Conservative",  293, False),
    ("Tipton Green", "Katarzyna Dobosz",         "Conservative",  261, False),
    ("Tipton Green", "Manjit Singh",             "Conservative",  247, False),
    ("Tipton Green", "Abdul Husen",              "Reform UK",     956, False),

    # Tividale
    ("Tividale", "Ken Parsons",                  "Reform UK",    1366, True),
    ("Tividale", "Lou Yates",                    "Reform UK",    1280, True),
    ("Tividale", "Maria Crompton",               "Labour Party", 1338, True),
    ("Tividale", "Coleen Sheehan",               "Reform UK",    1267, False),
    ("Tividale", "Altaf Hussain",                "Labour Party", 1025, False),
    ("Tividale", "Ritu Sharma",                  "Labour Party", 1022, False),
    ("Tividale", "Mark Byrne",                   "Green Party",   428, False),
    ("Tividale", "Robin Sarah Diver",            "Green Party",   381, False),
    ("Tividale", "James Francis Munoz",          "Green Party",   307, False),
    ("Tividale", "Jasbir Ranie",                 "Conservative",  253, False),
    ("Tividale", "Libbi Tudor",                  "Conservative",  297, False),
    ("Tividale", "Jacek Wojtowicz",              "Conservative",  214, False),
    ("Tividale", "Palwinder Singh",              "Lib Dem",       148, False),

    # Wednesbury
    ("Wednesbury", "Jeremy John Handley",        "Reform UK",    1777, True),
    ("Wednesbury", "Owen Michael Nelson",        "Reform UK",    1714, True),
    ("Wednesbury", "Paul Snape",                 "Reform UK",    1659, True),
    ("Wednesbury", "Mohammed All-Hasan",         "Green Party",   563, False),
    ("Wednesbury", "Eve Ward",                   "Green Party",   440, False),
    ("Wednesbury", "Jai Kaur Sandhu",            "Green Party",   411, False),
    ("Wednesbury", "Luke Giles",                 "Labour Party",  850, False),
    ("Wednesbury", "Peter Hughes",               "Labour Party",  823, False),
    ("Wednesbury", "Nicola Marie Maycock",       "Labour Party",  786, False),
    ("Wednesbury", "George Okpako",              "Conservative",  298, False),
    ("Wednesbury", "Evelyn Pessu",               "Conservative",  293, False),
    ("Wednesbury", "Rafiullah Mohammadzai",      "Conservative",  279, False),
    ("Wednesbury", "Richard Daniel Jones",       "Lib Dem",       206, False),
    ("Wednesbury", "Richard McVittie",           "Lib Dem",       188, False),
    ("Wednesbury", "Gina Patel",                 "Independent",   106, False),

    # West Bromwich Central
    ("West Bromwich Central", "Tirath Singh Dhatt",   "Labour Party", 1371, True),
    ("West Bromwich Central", "Liam Vincent Preece",  "Labour Party", 1221, True),
    ("West Bromwich Central", "Harpal Kaur Tiwana",   "Labour Party", 1197, True),
    ("West Bromwich Central", "Chris George",         "Reform UK",     586, False),
    ("West Bromwich Central", "Sue Taylor",           "Reform UK",     534, False),
    ("West Bromwich Central", "Gaynor Summan",        "Reform UK",     511, False),
    ("West Bromwich Central", "Oluwadamilola Amoo",   "Green Party",   483, False),
    ("West Bromwich Central", "Nikhwat Marawat",      "Green Party",   417, False),
    ("West Bromwich Central", "Kristopher Sarabadu",  "Green Party",   384, False),
    ("West Bromwich Central", "Sarah James",          "Conservative",  423, False),
    ("West Bromwich Central", "Suzanna Narojczyk",    "Conservative",  301, False),
    ("West Bromwich Central", "Anne-Marie Wright",    "Conservative",  319, False),
    ("West Bromwich Central", "Daria Gorczynska",     "Independent",   150, False),
]

WARD_META = {
    "Bearwood":                           {"electorate": 10385, "ballots": 4243, "turnout": 40.86},
    "Blackheath":                         {"electorate": 11227, "ballots": 3921, "turnout": 34.92},
    "Bristnall":                          {"electorate": 10529, "ballots": 3629, "turnout": 34.47},
    "Charlemont & Grove Vale":            {"electorate": 10202, "ballots": 4074, "turnout": 39.93},
    "Cradley Heath & Old Hill":           {"electorate": 10337, "ballots": 3358, "turnout": 32.49},
    "Friar Park & Stone Cross":           {"electorate": 10045, "ballots": 2919, "turnout": 29.60},
    "Great Barr, Tamebridge & Yew Tree":  {"electorate":  9832, "ballots": 3942, "turnout": 40.09},
    "Great Bridge":                       {"electorate": 10647, "ballots": 3460, "turnout": 32.50},
    "Greets Green & Lyng":                {"electorate":  9130, "ballots": 3246, "turnout": 35.55},
    "Hateley Heath":                      {"electorate":  9452, "ballots": 3056, "turnout": 32.33},
    "Hill Top":                           {"electorate":  9707, "ballots": 3281, "turnout": 33.80},
    "Langley":                            {"electorate": 10713, "ballots": 3551, "turnout": 33.15},
    "Newton & Valley":                    {"electorate":  9725, "ballots": 3544, "turnout": 36.44},
    "Old Warley":                         {"electorate":  9427, "ballots": 3735, "turnout": 39.62},
    "Oldbury":                            {"electorate":  8845, "ballots": 2850, "turnout": 32.22},
    "Princes End":                        {"electorate": 10045, "ballots": 2919, "turnout": 29.60},
    "Rowley":                             {"electorate": 10085, "ballots": 3426, "turnout": 33.97},
    "Smethwick":                          {"electorate":  9728, "ballots": 2993, "turnout": 30.77},
    "Soho & Victoria":                    {"electorate":  9174, "ballots": 2569, "turnout": 28.00},
    "St Paul's":                          {"electorate":  9421, "ballots": 3229, "turnout": 34.27},
    "Tipton Green":                       {"electorate": 10038, "ballots": 3886, "turnout": 38.71},
    "Tividale":                           {"electorate":  9592, "ballots": 3330, "turnout": 34.72},
    "Wednesbury":                         {"electorate": 10637, "ballots": 3644, "turnout": 34.26},
    "West Bromwich Central":              {"electorate":  9026, "ballots": 2880, "turnout": 31.91},
}

# ── PARTY GROUPINGS ─────────────────────────────────────────────────────────

def party_group(party):
    if party == "Labour Party":    return "Labour"
    if party == "Green Party":     return "Green"
    if party == "Reform UK":       return "Reform UK"
    if party == "Conservative":    return "Conservative"
    if party == "Lib Dem":         return "Lib Dem"
    if party == "Independent":     return "Independent"
    return "Other"

def bloc(party):
    g = party_group(party)
    if g in ("Labour", "Green"):   return "Left"
    if g in ("Reform UK", "Conservative"): return "Right"
    if g == "Lib Dem":             return "Lib Dem"
    return "Other/Independent"

# ── BUILD DATAFRAME ──────────────────────────────────────────────────────────

df = pd.DataFrame(WARDS, columns=["Ward", "Candidate", "Party", "Votes", "Elected"])
df["Party Group"] = df["Party"].apply(party_group)
df["Bloc"] = df["Party"].apply(bloc)
df["Elected"] = df["Elected"].map({True: "Yes", False: "No"})

# ── SHEET 1: RAW RESULTS ─────────────────────────────────────────────────────

raw = df[["Ward", "Candidate", "Party", "Votes", "Elected"]].copy()
raw = raw.sort_values(["Ward", "Votes"], ascending=[True, False]).reset_index(drop=True)

# ── SHEET 2: PARTY TOTALS ────────────────────────────────────────────────────

party_totals = (df.groupby("Party Group")["Votes"]
                  .sum()
                  .rename("Total Votes")
                  .reset_index()
                  .sort_values("Total Votes", ascending=False))
grand_total = party_totals["Total Votes"].sum()
party_totals["% of All Votes"] = (party_totals["Total Votes"] / grand_total * 100).round(2)

party_seats = (df[df["Elected"] == "Yes"]
               .groupby("Party Group")
               .size()
               .rename("Seats Won")
               .reset_index())
party_totals = party_totals.merge(party_seats, on="Party Group", how="left")
party_totals["Seats Won"] = party_totals["Seats Won"].fillna(0).astype(int)

# ── SHEET 3: BLOC ANALYSIS ────────────────────────────────────────────────────

bloc_totals = (df.groupby("Bloc")["Votes"]
                 .sum()
                 .rename("Total Votes")
                 .reset_index()
                 .sort_values("Total Votes", ascending=False))
bloc_totals["% of All Votes"] = (bloc_totals["Total Votes"] / grand_total * 100).round(2)

# ── SHEET 4: WARD-BY-WARD SPLIT VOTE ANALYSIS ────────────────────────────────

ward_records = []
for ward in sorted(df["Ward"].unique()):
    wdf = df[df["Ward"] == ward]
    meta = WARD_META.get(ward, {})

    reform_votes  = wdf[wdf["Party Group"] == "Reform UK"]["Votes"].sum()
    labour_votes  = wdf[wdf["Party Group"] == "Labour"]["Votes"].sum()
    green_votes   = wdf[wdf["Party Group"] == "Green"]["Votes"].sum()
    con_votes     = wdf[wdf["Party Group"] == "Conservative"]["Votes"].sum()
    libdem_votes  = wdf[wdf["Party Group"] == "Lib Dem"]["Votes"].sum()
    ind_votes     = wdf[wdf["Party Group"] == "Independent"]["Votes"].sum()
    other_votes   = wdf[wdf["Party Group"] == "Other"]["Votes"].sum()
    left_votes    = labour_votes + green_votes
    right_votes   = reform_votes + con_votes
    ward_total    = wdf["Votes"].sum()

    elected = wdf[wdf["Elected"] == "Yes"]["Party Group"].value_counts().to_dict()
    reform_seats  = elected.get("Reform UK", 0)
    labour_seats  = elected.get("Labour", 0)
    green_seats   = elected.get("Green", 0)
    con_seats     = elected.get("Conservative", 0)
    ind_seats     = elected.get("Independent", 0)

    # Top Reform candidate votes (the winner threshold)
    reform_top3 = wdf[wdf["Party Group"] == "Reform UK"]["Votes"].nlargest(3).tolist()
    reform_threshold = reform_top3[-1] if len(reform_top3) >= 3 else (reform_top3[-1] if reform_top3 else 0)

    # Left top3 combined votes
    left_candidates = wdf[wdf["Bloc"] == "Left"]["Votes"].nlargest(3).tolist()
    left_top3_sum = sum(left_candidates)

    # Split-vote indicator: did Reform win wards where Labour+Green > Reform?
    split_vote_evident = (left_votes > reform_votes) and (reform_seats > 0)

    ward_records.append({
        "Ward":                    ward,
        "Electorate":              meta.get("electorate", ""),
        "Ballots Cast":            meta.get("ballots", ""),
        "Turnout %":               meta.get("turnout", ""),
        "Reform UK Votes":         reform_votes,
        "Labour Votes":            labour_votes,
        "Green Votes":             green_votes,
        "Left (Lab+Green) Votes":  left_votes,
        "Conservative Votes":      con_votes,
        "Right (Ref+Con) Votes":   right_votes,
        "Lib Dem Votes":           libdem_votes,
        "Independent Votes":       ind_votes,
        "Other Votes":             other_votes,
        "Total Candidate Votes":   ward_total,
        "Reform % of Cand. Votes": round(reform_votes / ward_total * 100, 1) if ward_total else 0,
        "Labour % of Cand. Votes": round(labour_votes / ward_total * 100, 1) if ward_total else 0,
        "Green % of Cand. Votes":  round(green_votes  / ward_total * 100, 1) if ward_total else 0,
        "Left % of Cand. Votes":   round(left_votes   / ward_total * 100, 1) if ward_total else 0,
        "Right % of Cand. Votes":  round(right_votes  / ward_total * 100, 1) if ward_total else 0,
        "Reform Seats":            reform_seats,
        "Labour Seats":            labour_seats,
        "Green Seats":             green_seats,
        "Conservative Seats":      con_seats,
        "Independent Seats":       ind_seats,
        "Split Vote Evident?":     "YES" if split_vote_evident else "no",
        "Left Votes > Reform?":    "YES" if left_votes > reform_votes else "no",
        "Reform Win Threshold (3rd place votes)": reform_threshold,
    })

ward_df = pd.DataFrame(ward_records)

# ── SHEET 5: OVERALL LEFT/RIGHT PER WARD (for Flourish) ──────────────────────

flourish_df = ward_df[[
    "Ward",
    "Reform UK Votes", "Labour Votes", "Green Votes",
    "Conservative Votes", "Lib Dem Votes", "Independent Votes",
    "Left (Lab+Green) Votes", "Right (Ref+Con) Votes",
    "Reform % of Cand. Votes", "Labour % of Cand. Votes",
    "Green % of Cand. Votes", "Left % of Cand. Votes", "Right % of Cand. Votes",
    "Reform Seats", "Labour Seats", "Green Seats", "Conservative Seats",
    "Split Vote Evident?"
]].copy()

# ── WRITE EXCEL ───────────────────────────────────────────────────────────────

output_path = r"C:\Users\antoj\OneDrive\Escritorio\localelections_project\Sandwell\Sandwell_Elections_2026.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    raw.to_excel(writer, sheet_name="1. Raw Results", index=False)
    party_totals.to_excel(writer, sheet_name="2. Party Totals", index=False)
    bloc_totals.to_excel(writer, sheet_name="3. Left-Right Blocs", index=False)
    ward_df.to_excel(writer, sheet_name="4. Ward Analysis", index=False)
    flourish_df.to_excel(writer, sheet_name="5. Flourish Data", index=False)

# ── STYLE THE WORKBOOK ────────────────────────────────────────────────────────

wb = load_workbook(output_path)

PARTY_COLOURS = {
    "Reform UK":    "12B6CF",
    "Labour Party": "E4003B",
    "Labour":       "E4003B",
    "Green Party":  "02A95B",
    "Green":        "02A95B",
    "Conservative": "0087DC",
    "Lib Dem":      "FAA61A",
    "Independent":  "888888",
    "Other":        "AAAAAA",
}

HEADER_FILL  = PatternFill("solid", fgColor="1F3864")
HEADER_FONT  = Font(color="FFFFFF", bold=True)
ALT_FILL     = PatternFill("solid", fgColor="EBF1F8")
BORDER_SIDE  = Side(style="thin", color="CCCCCC")
THIN_BORDER  = Border(left=BORDER_SIDE, right=BORDER_SIDE,
                      top=BORDER_SIDE, bottom=BORDER_SIDE)

def style_sheet(ws, col_widths=None):
    # Header row
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER
    # Data rows
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        fill = ALT_FILL if i % 2 == 0 else PatternFill()
        for cell in row:
            cell.fill = fill
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
    # Column widths
    if col_widths:
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width
    else:
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)
    ws.freeze_panes = "A2"

# Sheet 1 – colour by party
ws1 = wb["1. Raw Results"]
style_sheet(ws1)
party_col = {cell.value: cell.column for cell in ws1[1] if cell.value == "Party"}
for row in ws1.iter_rows(min_row=2):
    party_val = row[2].value  # Party column is index 2
    colour = PARTY_COLOURS.get(party_val)
    if colour:
        row[2].fill = PatternFill("solid", fgColor=colour)
        row[2].font = Font(color="FFFFFF", bold=True)
    elected_cell = row[4]
    if elected_cell.value == "Yes":
        elected_cell.fill = PatternFill("solid", fgColor="00B050")
        elected_cell.font = Font(color="FFFFFF", bold=True)

# Sheet 2 – party totals
ws2 = wb["2. Party Totals"]
style_sheet(ws2)
for row in ws2.iter_rows(min_row=2):
    party_val = row[0].value
    colour = PARTY_COLOURS.get(party_val)
    if colour:
        row[0].fill = PatternFill("solid", fgColor=colour)
        row[0].font = Font(color="FFFFFF", bold=True)
    pct_cell = row[2]
    if pct_cell.value:
        pct_cell.number_format = "0.00%"

# Sheet 3 – blocs
ws3 = wb["3. Left-Right Blocs"]
style_sheet(ws3)
BLOC_COLOURS = {"Left": "E4003B", "Right": "12B6CF", "Lib Dem": "FAA61A", "Other/Independent": "888888"}
for row in ws3.iter_rows(min_row=2):
    bloc_val = row[0].value
    colour = BLOC_COLOURS.get(bloc_val)
    if colour:
        row[0].fill = PatternFill("solid", fgColor=colour)
        row[0].font = Font(color="FFFFFF", bold=True)

# Sheet 4 – ward analysis, highlight split vote
ws4 = wb["4. Ward Analysis"]
style_sheet(ws4)
split_col_idx = None
for cell in ws4[1]:
    if cell.value == "Split Vote Evident?":
        split_col_idx = cell.column
        break
if split_col_idx:
    for row in ws4.iter_rows(min_row=2):
        cell = row[split_col_idx - 1]
        if cell.value == "YES":
            cell.fill = PatternFill("solid", fgColor="FF0000")
            cell.font = Font(color="FFFFFF", bold=True)

# Sheet 5 – flourish data
ws5 = wb["5. Flourish Data"]
style_sheet(ws5)

wb.save(output_path)
print(f"Saved: {output_path}")

# ── SUMMARY PRINTOUT ──────────────────────────────────────────────────────────
print("\n=== OVERALL PARTY VOTE SHARES ===")
print(party_totals.to_string(index=False))
print("\n=== LEFT/RIGHT BLOCS ===")
print(bloc_totals.to_string(index=False))
print("\n=== WARDS WHERE SPLIT VOTE IS EVIDENT (Left > Reform AND Reform won seats) ===")
split_wards = ward_df[ward_df["Split Vote Evident?"] == "YES"][
    ["Ward", "Reform UK Votes", "Left (Lab+Green) Votes", "Reform Seats", "Labour Seats", "Green Seats"]
]
print(split_wards.to_string(index=False))
print(f"\nTotal wards with split vote evidence: {len(split_wards)} / {len(ward_df)}")
