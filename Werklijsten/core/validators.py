import os
import magic
from django.core.exceptions import ValidationError


def validate_is_xlsx(file):
    valid_mime_types = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
    file_mime_type = magic.from_buffer(file.read(2048), mime=True)
    if file_mime_type not in valid_mime_types:
        print(file_mime_type)
        raise ValidationError('Unsupported file type.')
    valid_file_extensions = ['.xlsx']
    ext = os.path.splitext(file.name)[1]
    if ext.lower() not in valid_file_extensions:
        raise ValidationError('Unacceptable file extension.')
