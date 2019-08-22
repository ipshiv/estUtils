# -*- coding: utf-8 -*-

from app import logger, db, utils
from app.models import File, Task

from arps import generate_arps_from_statement
from matches import detect_matches
from smeta import generate_arps_from_smeta
from common_type import common_type_xls_links

class ProcessorNotFound(Exception):
    pass

def execute(task):

    queue_name = task.queue_name

    if queue_name == 'arps':
        run_processor = generate_arps_from_statement
    elif queue_name == 'smeta':
        run_processor = generate_arps_from_smeta
    elif queue_name == 'common_type':
        run_processor = common_type_xls_links
    elif queue_name == 'matches':
        run_processor = detect_matches
    else:
        logger.debug('Can\'t find processor \'%s\'', name)
        raise ProcessorNotFound('Can\'t find processor \'%s\'' % name)

    try:
        if queue_name in ['arps', 'smeta', 'common_type']:
            try:
                source = next(source for source in task.sources if source.type == 'source')
                result_path = run_processor(source.path)
            except StopIteration:
                raise Exception('Illegal state of task %d, can\'t find record about source files', task_id)
        elif queue_name == 'matches':
            try:
                source = next(source for source in task.sources if source.type == 'source')
                template = next(source for source in task.sources if source.type == 'template')
                result_path = run_processor(source.path, template.path, template.filename, task.threshold)
            except StopIteration:
                raise Exception('Illegal state of task %d, can\'t find record about source files', task_id)

        if result_path != None and len(result_path) > 0:
            filename = utils.extract_file_name(result_path)
            result = File(filename, result_path)
            result.type = File.RESULT
            task.result = result
            task.status = Task.COMPLETED
        else:
            task.status = Task.FAILED
    except Exception, e:
        task.status = Task.FAILED
        raise e
    finally:
        db.session.commit()

