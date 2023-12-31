summary_reference_dict = {'params': {
    "shroud": {
        "stg_1": "B6:H9",
        "stg_2": "B11:H14",
        "stg_3": "B16:H19",
        "stg_4": "B21:H24",
        "stg_4_ref": "B26:H29"},
    "nozzle": {
        "stg_2": "B32:H37",
        "stg_3": "B39:H44",
        "stg_4": "B46:H51"},
},
    'file_name': 'summary.json'}

judgment_reference_dict = {'params': {
    "shroud": {
        "stg_1": "C11:H15",
        "stg_2": "C55:H59",
        "stg_3": "C99:H103",
        "stg_4": "C144:H148",
        "stg_4_ref": "C187:H191"},
    "nozzle": {
        "stg_2": "C232:H236",
        "stg_3": "C276:H280",
        "stg_4": "C320:H324"},
},
    'file_name': 'judgment.json'}

radius_reference_dict = {'params': {
    "radius": {
        "stg_gen": "AB55:AF63",},
},
    'file_name': 'radius.json'}


slide_indices = {"shroud": [6,7,8,9,10], "nozzle": [14,15,16], "radius": [35]}

dimensions_dict = {
    "image_type_1": {"image_height": 11.2,
                     "image_width": 14.85,
                     "horizontal_pos": 0.9,
                     "vertical_pos": 6.93},
    "image_type_2": {"image_height": 6.0,
                     "image_width": 7.95,
                     "horizontal_pos": 16.5,
                     "vertical_pos": 6.53},

    "image_type_2a": {"image_height": 2.75,
                     "image_width": 15,
                     "horizontal_pos": 9,
                     "vertical_pos": 2.1},

    "image_type_2b": {"image_height": 3.2,
                     "image_width": 8.5,
                     "horizontal_pos": 16.5,
                     "vertical_pos": 14.5},

    # slide12
    "image_type_3": {"image_height": 7.92,
                     "image_width": 11.27,
                     "horizontal_pos": 0.94,
                     "vertical_pos": 7.3},
    "image_type_4": {"image_height": 7.92,
                     "image_width": 11.27,
                     "horizontal_pos": 12.89,
                     "vertical_pos": 7.3},
    # slide13
    "image_type_5": {"image_height": 7.92,
                     "image_width": 11.65,
                     "horizontal_pos": 0.9,
                     "vertical_pos": 7.3},
    "image_type_6": {"image_height": 7.92,
                     "image_width": 11.65,
                     "horizontal_pos": 13.07,
                     "vertical_pos": 7.31},
    # slide14
    "image_type_7": {"image_height": 10.28,
                     "image_width": 12.15,
                     "horizontal_pos": 0.5,
                     "vertical_pos": 4.56},
    "image_type_8": {"image_height": 10.28,
                     "image_width": 12.15,
                     "horizontal_pos": 12.98,
                     "vertical_pos": 4.52},
    # slide36
    "image_type_9": {"image_height": 11.2,
                     "image_width": 16.15,
                     "horizontal_pos": 0.9,
                     "vertical_pos": 6.93},
    # slide 19 20 21 23 29 30 31 32 34
    "image_type_10": {"image_height": 12.6,
                      "image_width": 12.25,
                      "horizontal_pos": 0.5,
                      "vertical_pos": 3.64},
    "image_type_11": {"image_height": 12.28,
                      "image_width": 12.15,
                      "horizontal_pos": 12.9,
                      "vertical_pos": 3.65},
    # slide 22
    "image_type_12": {"image_height": 12.6,
                      "image_width": 12.16,
                      "horizontal_pos": 6.62,
                      "vertical_pos": 3.61},
    # slide 26 27
    "image_type_13": {"image_height": 12.6,
                      "image_width": 16.77,
                      "horizontal_pos": 3.9,
                      "vertical_pos": 3.72},

    # slide 36
    "image_type_14": {"image_height": 11.2,
                      "image_width": 16.15,
                      "horizontal_pos": 0.9,
                      "vertical_pos": 6.93},

}
image_filename_slide_dict = {
    "Graph_9.png": {"slide_index": 6,
                    "image_type": "image_type_1"},
    "Graph_0.png": {"slide_index": 6,
                    "image_type": "image_type_2"},

    "Graph_10.png": {"slide_index": 7,
                     "image_type": "image_type_1"},
    "Graph_1.png": {"slide_index": 7,
                    "image_type": "image_type_2"},
    
    "Graph_11.png": {"slide_index": 8,
                     "image_type": "image_type_1"},
    "Graph_2.png": {"slide_index": 8,
                    "image_type": "image_type_2"},
    
    "Graph_15.png": {"slide_index": 9,
                     "image_type": "image_type_1"},
    "Graph_8.png": {"slide_index": 9,
                    "image_type": "image_type_2"},
    
    "Graph_16.png": {"slide_index": 10,
                     "image_type": "image_type_1"},
    "Graph_3.png": {"slide_index": 10,
                    "image_type": "image_type_2"},
    
    "246_0.png": {"slide_index": 11,
                  "image_type": "image_type_3"},
    "246_1.png": {"slide_index": 11,
                  "image_type": "image_type_4"},
    "201_0.png": {"slide_index": 12,
                  "image_type": "image_type_5"},
    "201_1.png": {"slide_index": 12,
                  "image_type": "image_type_6"},
    "201-Tool_max_2.png": {"slide_index": 13,
                           "image_type": "image_type_7"},
    "201-Tool_min_2.png": {"slide_index": 13,
                           "image_type": "image_type_8"},
    "Graph_12.png": {"slide_index": 14,
                     "image_type": "image_type_1"},
    "Graph_4.png": {"slide_index": 14,
                    "image_type": "image_type_2"},
    
    "Graph_13.png": {"slide_index": 15,
                     "image_type": "image_type_1"},
    "Graph_5.png": {"slide_index": 15,
                    "image_type": "image_type_2"},
    
    "Graph_14.png": {"slide_index": 16,
                     "image_type": "image_type_1"},
    "Graph_6.png": {"slide_index": 16,
                    "image_type": "image_type_2"},
    

    # part2 - needs to be updated from here

    "105_0.png": {"slide_index": 19,
                  "image_type": "image_type_10"},
    "234_0.png": {"slide_index": 19,
                  "image_type": "image_type_11"},

    "235_0.png": {"slide_index": 20,
                  "image_type": "image_type_10"},
    "235_1.png": {"slide_index": 20,
                  "image_type": "image_type_11"},

    "234-235Stg1SHookT_0.png": {"slide_index": 21,
                                "image_type": "image_type_12"},

    "105-235Stg1SGroove_0.png": {"slide_index": 22,
                                  "image_type": "image_type_10"},
    "105-235Stg1SGroove_1.png": {"slide_index": 22,
                                   "image_type": "image_type_11"},

    "112_0.png": {"slide_index": 25,
                  "image_type": "image_type_13"},

    "118_0.png": {"slide_index": 26,
                  "image_type": "image_type_13"},

    "207_1.png": {"slide_index": 28,
                  "image_type": "image_type_10"},
    "207_2.png": {"slide_index": 28,
                  "image_type": "image_type_11"},
    "225_1.png": {"slide_index": 29,
                  "image_type": "image_type_10"},
    "225_2.png": {"slide_index": 29,
                  "image_type": "image_type_11"},
    "226_1.png": {"slide_index": 30,
                  "image_type": "image_type_10"},
    "226_2.png": {"slide_index": 30,
                  "image_type": "image_type_11"},
    # "225226CaseThickness_0.png": {"slide_index": 31,
    #                               "image_type": "image_type_10"},
    "225226CaseThickness_0.png": {"slide_index": 31,
                                   "image_type": "image_type_11"},
    "101_0.png": {"slide_index": 33,
                  "image_type": "image_type_10"},
    "101_1.png": {"slide_index": 33,
                  "image_type": "image_type_11"},
    "Graph_7.png": {"slide_index": 35,

                  "image_type": "image_type_14"},
    

}

