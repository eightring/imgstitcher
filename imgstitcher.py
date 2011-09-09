#!/usr/bin/python
# encoding: utf-8

"""image stitcher - combine mutiple images into one"""
__version__ = "3.14"
__author__  = 'yjqian (eight.ring@gmail.com)'

import os, sys, getopt, time, subprocess
from PIL import Image

g_selfname = os.path.realpath(sys.argv[0])
g_ppt2img  = os.path.join(os.path.dirname(g_selfname), 'ppt2img2.vbs')

def isfile(filepath, types):
    if not filepath:
        return False
    if os.path.isdir(filepath):
        return False
    ext = filepath.rsplit('.', 1)[-1].lower()
    if ext in types:
        return True
    return False

def isimage(filepath):
    return isfile(filepath, ['bmp', 'gif', 'jpeg', 'jpg', 'png'])
  
def isppt(filepath):
    return isfile(filepath, ['ppt', 'pptx'])

def get_images(dirpath):
    imglist = []
    for sub in os.listdir(dirpath):
        fullpath = os.path.join(dirpath, sub)
        if os.path.isdir(fullpath):
            continue
        if isimage(fullpath):
            imglist.append(fullpath)
    return imglist

def run_command(cmd):
    "run a command"
    try:
        retcode = subprocess.call(cmd, shell=True)
    except OSError, e:
        print >> sys.stderr, "run command %s failed: %s" % (cmd, e)
    finally:
        return retcode

def resize(im, percent):
    """ shrink according to a percentage """
    w, h = im.size
    return im.resize(((percent*w)/100,(percent*h)/100))

def resize2(im, maxw):
    """ resize the image, so that width no big than maxw """
    w, h = im.size
    rw = 1.0 * w / maxw
    return im.resize((int(w/rw), int(h/rw)))

def resize3(im, maxh):
    """ resize the image, so that heightth no big than maxh """
    w, h = im.size
    ratio = 1.0 * h / maxh
    return im.resize((int(w/ratio), int(h/ratio)))

def resize4(im, pixels):
    """ resize the longest side in pixels """
    (wx, wy) = im.size
    rx = 1.0 * wx / pixels
    ry = 1.0 * wy / pixels
    if rx > ry:
        rr = rx
    else:
        rr = ry
    return im.resize((int(wx/rr), int(wy/rr)))

def print_version():
    print __doc__
    print 'version', __version__
    print 'author', __author__

def usage(msg = None):
    if msg: print msg
    prog = os.path.basename(sys.argv[0])
    print prog, "[-h] [-v] file1 file2 file3 ..."
    print "  -h         print usage"
    print "  -v         print version"
    print "  -s space   sapce between images. default 5px"
    print

g_max_width  = 800
g_img_width  = 0
g_img_height = 0
g_img_space  = 5
g_img_list   = []

def main():
    global g_img_width, g_img_height, g_img_space, g_img_list

    try: opts, args = getopt.getopt(sys.argv[1:], "hvs:", [])
    except getopt.GetoptError, e:
        usage(str(e))
        return 1

    for opt, value in opts:
        if opt == "-s":
            g_img_space = int(value)
        if opt == "-h":
            usage()
            return 0
        if opt == "-v":
            print_version()
            return 0

    if len(args) == 0:
        usage()
        return 1

    if len(args) == 1 and isppt(args[0]):
        run_command('wscript "%s" "%s"' % (g_ppt2img, args[0]))
        imgs = get_images(args[0].rsplit('.', 1)[0])
        g_img_list.extend(imgs)
    elif len(args) == 1 and os.path.isdir(args[0]):
        g_img_list.extend(get_images(args[0]))
    else:
        for x in args:
            if isimage(x):
                g_img_list.append(x)

    g_img_list.sort()

    result_dir = os.path.dirname(args[0])
    
    for x in g_img_list:
        try:
            img = Image.open(x)
            w, h = img.size
            if w > g_max_width:
                h = int(1.0 * h * g_max_width / w)
                w = g_max_width
            if w > g_img_width:
                g_img_width = w
            g_img_height += h + g_img_space
        except Exception, err:
            print >>sys.stderr, err
            
    final_img = Image.new("RGB", (g_img_width, g_img_height))
    offset = 0
    
    for x in g_img_list:
        tmp = Image.open(x)
        if tmp.size[0] > g_max_width:
            tmp = resize2(tmp, g_max_width)
        w, h = tmp.size
        final_img.paste(tmp, (int((g_img_width-w)/2), offset))
        offset += tmp.size[1] + g_img_space
    
    final_img.save(os.path.join(result_dir, 'final_%s.png' % time.time()))

    return 0

if __name__ == "__main__":
    sys.exit(main())
