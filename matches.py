# -*- coding: utf-8 -*-

import os
import random
import string

import cv2

from app import app, logger, utils

DEFAULT_THRESHOLD = 0.7

def detect_matches(source_path, template_path, template_name, threshold=DEFAULT_THRESHOLD):
    image = _load_image(source_path)
    template = _load_image(template_path)

    if threshold == None:
        threshold = DEFAULT_THRESHOLD

    found_matches, result = _find_matches(image, template, threshold)

    text = 'Template %s found on image: %d times!' % (template_name, found_matches)
    font = cv2.FONT_HERSHEY_SIMPLEX
    cv2.putText(result, text.encode('utf-8', errors='ignore'), (10, 200), font, 4, (0, 0, 255), 2)

    image_filename_parts = source_path.split('/')[-1].split('.')
    image_ext = image_filename_parts[1] if len(image_filename_parts) >= 2 else 'png'
    result_filename = utils.generate_random_string(length=10)
    result_path = app.config['PROCESSING_RESULTS_DIR'] + '/%s.%s' % (result_filename, image_ext)
    cv2.imwrite(result_path, result)

    return result_path

def _find_matches(image, template, threshold):
    gry_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gry_template = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)

    _, thr_image = cv2.threshold(gry_image, 127, 255, cv2.THRESH_BINARY)
    _, thr_template = cv2.threshold(gry_template, 127, 255, cv2.THRESH_BINARY)

    objects = []

    rtd_templates = _get_rotated_template_images(thr_template)
    for rtd_template in rtd_templates:
        height, weight = rtd_template.shape[:2]

        result = cv2.matchTemplate(thr_image, rtd_template, cv2.TM_CCOEFF_NORMED)

        while(True):
            # Повторяем цикл пока в матрице result есть совпадения
            (_, max_val, _, max_loc) = cv2.minMaxLoc(result)
            if max_val >= threshold:
                # Удаляем область содержащую найденный объект из матрицы result
                cv2.rectangle(result, (max_loc[0] - weight / 2, max_loc[1] - height / 2), \
                    (max_loc[0] + weight / 2, max_loc[1] + height / 2), (0, 0, 0), -1)

                # Исключаем дубликаты
                top_left_point = max_loc
                bottom_right_point = (max_loc[0] + weight, max_loc[1] + height)

                exists = False
                for obj in objects:
                    x1 = abs(top_left_point[0] - obj['top_left_point'][0])
                    y1 = abs(top_left_point[1] - obj['top_left_point'][1])
                    x2 = abs(bottom_right_point[0] - obj['bottom_right_point'][0])
                    y2 = abs(bottom_right_point[1] - obj['bottom_right_point'][1])

                    if (x1 <= weight and y1 <= height) \
                        or (x2 <= weight and y2 <= height):
                       exists = True
                       break

                if not exists:
                    objects.append({
                        'top_left_point': top_left_point,
                        'bottom_right_point': bottom_right_point
                    })

                    cv2.rectangle(image, max_loc, \
                        (max_loc[0] + weight, max_loc[1] + height), (0, 255, 0), 2)
            else:
                break

    return len(objects), image

def _get_rotated_template_images(template_image):
    return [_rotate_image(template_image, angle)
        for angle in (45, 90, 135, 180, 225, 270, 315, 360)]

def _rotate_image(image, angle):
    rows, cols = image.shape[:2]
    rotation_matrix = cv2.getRotationMatrix2D((cols / 2, rows / 2), angle, 1)

    return cv2.warpAffine(image, rotation_matrix, (cols, rows))

def _load_image(path):
    return cv2.imread(path)

def _save_image(name, image):

    cv2.imwrite(app.config['MATCHES_PROCESSOR_TMP_DIR'] + '/' + name, image)
