# Empty Public Folder Cleanup Process

## Removal Process

- Import the approved list of folders to be processed for potential removal (entryIDs)
- Submit the list to the processing function to perform validation(s)
  - get current public folder object
    - verify no subfolders
    - verify not mail enabled
  - get current public folder stats from each replica (from the public folder object)
    - verify no items
- validate
  - true, log the validation
    - submit the public folder entryid to the function for public folder removal
  - false, log the validation failure(s)

## Restore Process

- Import the approved list of folders to be processed for restoral
- Import the list of applicable public folder permissions
- submit the list(s) to the processing functions to restore the public folder attributes and permissions
- NOTE: this does not include any data restoral if data was deleted by the previous removal of a public folder
