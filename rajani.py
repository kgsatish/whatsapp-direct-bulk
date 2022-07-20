import configparser
import os
import logging
import config

cfg = configparser.ConfigParser()
cfg.read('rajani.ini')

excel_file = cfg['default']['excel_file']
img_files = cfg['default']['img_files'].split(',')
msg_files = cfg['default']['msg_files'].split(',')

logging.info('Image info: [%d] %s', len(img_files), img_files)
logging.info('Message info: [%d] %s', len(msg_files), msg_files)

if len(msg_files) == 1:
    logging.info('Inside one')
    level = 'A'
    img_file = ''
    if len(img_files) == 1:
        img_file = img_files[0].strip()
    msg_file = msg_files[0].strip()
    logging.info('Message [%s] [%s]', msg_file, img_file)
    if img_file:
        cmd = 'python app.py --img ' + img_file + ' ' + excel_file + ' ' + msg_file + ' ' + level
    else:
        cmd = 'python app.py ' + excel_file + ' ' + msg_file + ' ' + level
    logging.info('os.system(%s)', cmd)
    os.system(cmd)
    exit()
    
for idx in range(0, len(msg_files)):
    logging.info('Inside loop')
    img_file = ''
    if idx < len(img_files):
        img_file = img_files[idx].strip()
    msg_file = msg_files[idx].strip()
    level = str(idx+1)
    logging.info('Message [%d] [%s] [%s]', idx, msg_file, img_file)
    
    if img_file:
        cmd = 'python app.py --img ' + img_file + ' ' + excel_file + ' ' + msg_file + ' ' + level
    else:
        cmd = 'python app.py ' + excel_file + ' ' + msg_file + ' ' + str(idx+1)
    logging.info('os.system(%s)', cmd)
    os.system(cmd)
logging.info('Done with everything')
exit()