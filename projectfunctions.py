def pt_in_cm(unit_before_transform, after_transform):
    #磅值、英寸、厘米三者转换，参数为转换前后的单位种类，返回直接可乘的系数
    point = 72.0
    cm = 2.54
    pt_in_cm = 1
    if unit_before_transform == 1:
        pt_in_cm /= point
    elif unit_before_transform == 3:
        pt_in_cm /= cm
    if after_transform == 1:
        pt_in_cm *= point
    elif after_transform == 3:
        pt_in_cm *= cm
    return pt_in_cm

def a4_paper_height():
    # return width in cm
    return 29.7

def a4_paper_width():
    # return length in cm
    return 21.0
