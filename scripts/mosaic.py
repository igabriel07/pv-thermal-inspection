import glob, rasterio

paths = sorted(glob.glob(r"F:\project\thermal1_images\tiff\*.tif*"))[:20]
for p in paths:
    with rasterio.open(p) as src:
        print(p.split("\\")[-1], src.crs)
