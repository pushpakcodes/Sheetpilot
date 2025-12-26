import React, { useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import { Upload } from 'lucide-react';
import { cn } from '../utils';

const FileUploader = ({ onUpload, className }) => {
  const onDrop = useCallback((acceptedFiles) => {
    if (acceptedFiles?.length > 0) {
      onUpload(acceptedFiles[0]);
    }
  }, [onUpload]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    maxFiles: 1
  });

  return (
    <div
      {...getRootProps()}
      className={cn(
        'border-2 border-dashed rounded-xl p-8 text-center cursor-pointer transition-colors',
        isDragActive
          ? 'border-blue-500 bg-blue-500/10'
          : 'border-slate-700 hover:border-blue-500 hover:bg-slate-800/50',
        className
      )}
    >
      <input {...getInputProps()} />
      <Upload className="mx-auto h-12 w-12 text-slate-400 mb-4" />
      <p className="text-lg font-medium text-slate-200">
        {isDragActive ? 'Drop the file here' : 'Drag & drop Excel file here'}
      </p>
      <p className="text-sm text-slate-400 mt-2">
        or click to select file
      </p>
    </div>
  );
};

export default FileUploader;
