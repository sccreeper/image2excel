from lib import image2excel as i2e

data = []

add = []

for i in range(256):
	add.append((255,255,255))

for i in range(256):
	data.append(add)

converter = i2e.ImageConverter(data, "array_test")

converter.convert()