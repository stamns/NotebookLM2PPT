# from pyinpaint import Inpaint
import numpy as np
from skimage.restoration import inpaint
from PIL import Image
from .edge_diversity import compute_edge_diversity_numpy


def inpaint_image(image_path, output_path):
    # 2867, 1600
    image = Image.open(image_path)
    image_defect = np.array(image)
    mask = np.zeros(image_defect.shape[:-1], dtype=bool)
    # [{\"width\":240,\"top\":1530,\"height\":65,\"left\":2620}]
    r1,r2,c1,c2 = 1530,1595,2620,2860

    old_width, old_height = 2867,1600

    image_width, image_height = image_defect.shape[1], image_defect.shape[0]
    ratio = image_width / old_width

    assert abs(ratio - (image_height / old_height)) < 0.01, "图片比例不对，无法修复"


    r1 = int(r1 * ratio)
    r2 = int(r2 * ratio)
    c1 = int(c1 * ratio)
    c2 = int(c2 * ratio)

    edge_diversity, fill_color = compute_edge_diversity_numpy(image_defect, c1, r1, c2, r2)

    if edge_diversity < 0.1: # 直接填充完事
        print("直接填充",edge_diversity, fill_color)
        image_defect[r1:r2, c1:c2] = fill_color
        image_result = image_defect
    else:
        print("需要修复",edge_diversity, fill_color)
        mask[r1:r2, c1:c2] = True
        image_result = inpaint.inpaint_biharmonic(image_defect, mask, channel_axis=-1)
        image_result = (image_result*255).astype("uint8")
    Image.fromarray(image_result).save(output_path)
