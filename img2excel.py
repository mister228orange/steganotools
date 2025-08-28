import xlsxwriter
from PIL import Image

MAGIC_CONSTS = {
	"height": 374, # because it is count of cells who placed in my monitor in same column
	"zoom": 10 # because its minimal zoom which opened on my pc my excel my os
	# you may change this to your needs its not Physical constants
}

def transform(pixel):
	# 3 num to excel format
	return ''.join([f'{c:02x}' for c in pixel])

def img2pixels(img):
	#pixel_image = img.quantize(colors=256, method=Image.FASTOCTREE)
	rgb_im = img.convert('RGB')
	size = rgb_im.size
	height = MAGIC_CONSTS["height"]
	k = size[1] / height
	if k < 1:
		k = 1
		height = size[1]
	width = int(size[0] // k)
	small_img = rgb_im.resize((width, height), Image.BILINEAR)
	small_img.save("2step.png")
	pixels_list = [pixel for pixel in list(small_img.getdata())]
	pixels = [pixels_list[start * width: start * width + width] for start in range(height)]
	return pixels

def add_bullshit_legend(worksheet):
	pass

def save_as_excel(file_name, pixels, stealth_mode=1):
	workbook = xlsxwriter.Workbook(file_name)
	worksheet = workbook.add_worksheet()

	worksheet.set_column_pixels(0, len(pixels[0]), 10)
	worksheet.zoom = 10

	for i, row in enumerate(pixels):
		for j, pixel in enumerate(row):
			color = transform(pixel)
			if not stealth_mode:
				cell_format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': color})
				worksheet.write(i, j, '', cell_format)
			else:
				worksheet.write(i, j, int(pixel, 16))
	if stealth_mode > 1:
		add_bullshit_legend(worksheet)
	workbook.close()


def main():
	img = Image.open('HIGH.png')
	pixels = img2pixels(img)
	save_as_excel('Годовой отчет some boring stuff by shtribans.xlsx', pixels, 0)

#print(transform((10, 8, 20)))
main()