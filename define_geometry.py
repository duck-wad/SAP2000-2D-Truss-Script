def generate_warren(height, module_length, module_divisions, segment_length, num_modules):

    bottom_chord_points = []
    top_chord_points = []
    diagonal_web_points = []
    vertical_web_points = []

    # generate bottom chord points for first module
    for i in range(module_divisions + 1):
        bottom_chord_points.append((i*segment_length, 0.0, 0.0))
    # copy the chord over for the other modules
    # exclude the first point since it's already accounted for
    temp = len(bottom_chord_points) - 1
    for i in range(num_modules - 1):
        for j in range(temp):
            bottom_chord_points.append((bottom_chord_points[j+1][0] + (i+1)*module_length, 0.0, 0.0))

    # generate top chord points for first module
    for i in range(module_divisions + 2):
        if i == 0:
            top_chord_points.append((0.0, 0.0, height))
        elif i == module_divisions + 1:
            top_chord_points.append((module_length, 0.0, height))
        else:
            top_chord_points.append((i*segment_length - 0.5*segment_length, 0.0, height))
    # copy chord over for the other modules
    temp = len(top_chord_points) - 1
    for i in range(num_modules - 1):
        for j in range(temp):
            top_chord_points.append((top_chord_points[j+1][0] + (i+1)*module_length, 0.0, height))
    
    # generate the vertical web points which occur at x locations of multiples of module_length
    # for num_modules, there will be num_modules+1 vertical webs
    # in order (bottom_1, top_1, bottom_2, top_2, ....)
    for i in range(num_modules + 1):
        vertical_web_points.append((i*module_length, 0.0, 0.0))
        vertical_web_points.append((i*module_length, 0.0, height))

    # generate diagonal web points for first module
    bottom_counter = 0
    top_counter = 1
    for i in range(2 * module_divisions + 1):
        if i % 2 == 0:
            diagonal_web_points.append(bottom_chord_points[bottom_counter])
            bottom_counter += 1
        else:
            diagonal_web_points.append(top_chord_points[top_counter])
            top_counter += 1
    # copy webs over for the other modules
    temp = len(diagonal_web_points) - 1
    for i in range(num_modules - 1):
        for j in range(temp):
            diagonal_web_points.append((diagonal_web_points[j+1][0] + (i+1)*module_length, 0.0, 
                                        diagonal_web_points[j+1][2]))
    return bottom_chord_points, top_chord_points, diagonal_web_points, vertical_web_points