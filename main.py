from pptx import Presentation
import argparse
import sys, time

def generator(in_file, out_file):

    prs = Presentation(in_file)
    horizontal, column, images = get_setting(prs)

    # 1. 가로 세로를 순서대로 정렬하는 코드
    horizontal = sorted(horizontal, key=lambda x: int(x.top))
    column = sorted(column, key=lambda x: int(x.left))

    # 2. 가로의 개수만큼씩 나누는 코드.
    images = sorted(images, key=lambda x: int(x.top))
    temp = [[] for _ in range(len(column))]

    for i in range(len(column)):
        temp[i] = images[i*len(horizontal): (i+1)*len(horizontal)]
        temp[i] = sorted(temp[i], key=lambda x: int(x.left))
    images = temp

    # 3. align
    for i in range(len(column)):
        for j in range(len(horizontal)):
            images[i][j].top = column[i].top
            images[i][j].left = horizontal[j].left

    prs.save(out_file)

def get_setting(prs):
    shapes = prs.slides[0].shapes
    horizontal = []
    column = []
    images = []
    for shape in shapes:
        if int(shape.left) < 0:
            column.append(shape)
        elif int(shape.top) < 0:
            horizontal.append(shape)
        else:
            images.append(shape)

    # number check
    if (len(images) != len(column) * len(horizontal)):
        print("The number of images is not same with grid shape")
        time.sleep(3)
        sys.exit()

    return horizontal, column, images

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--infile', type=str, help='input file path')
    parser.add_argument('--outfile', type=str, help='output file path')
    args = parser.parse_args()

    generator(args.infile, args.outfile)