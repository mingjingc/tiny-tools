import glob
import rawpy
import time
import os
from PIL import Image
import shutil
from pathlib import Path
import io


def main():
    baseDir = Path("/*wrapper dir of nef images dirs*/")
    for basePathNEF in baseDir.iterdir():
        if basePathNEF.is_dir():
            # Timer is optional
            print("Starting processing:", os.path.basename(basePathNEF))
            start = time.time()

            # # Sometimes the file ending can be .NEF, therefore I included both possibilities to save extra work.
            dirNEF = sprint(basePathNEF.absolute(), "/*.NEF")
            dirResult = sprint(
                basePathNEF, '/', os.path.basename(basePathNEF))
            if not os.path.exists(dirResult):
                os.makedirs(dirResult)

            count = 0
            for path in glob.glob(dirNEF):
                with rawpy.imread(path) as raw:
                    rgb = raw.postprocess(use_camera_wb=True)
                    path = os.path.basename(path)
                    filename = path.split('.')[0]
                    imagePath = dirResult + '/' + filename + '.jpg'
                    Image.fromarray(rgb).save(imagePath,
                                              quality=100, optimize=True)
                    count = count + 1
                    print('image #', count, ' done')

            shutil.make_archive(dirResult, 'zip', dirResult)

            # calculate time spend
            end = time.time()
            minutes = (end - start) / 60
            seconds = (end - start) % 60
            print("Elapsed Time:", round(minutes),
                  "minutes and ", round(seconds), "seconds")
            print('======================================')


def sprint(*args, **kwargs):
    sio = io.StringIO()
    print(*args, **kwargs, sep='', end='', file=sio)
    return sio.getvalue()


if __name__ == '__main__':
    main()
