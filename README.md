# YoutubeToPPTX
Convert any youtube video into a pptx presentation

## Usage
### Input Format

The script can be used with any text file as long as the content complies with the format:

1. Line must be the ID of the Youtube video

All other lines must have the format:
`timestamp command [args]`

- Timestamps must be in the format HH:MM:SS or HH:MM:SS.FRACTION
- possible commands:
    - `\copy` copy clip from current to next time stamp into presentation
    - `\note` same as `\copy` but content of args is also coped into the slide notes
    - `\skip` skips to next command
    - undefined commands are handled as `\skip`

Comments can be added with lines starting with a #.

### Arguments
```
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        Presentation Manifest used to generate presentation
  -t TEMPLATE, --template TEMPLATE
                        Template powerpoint file, the generated presentation will be appended to the file
  -f FORMAT, --format FORMAT
                        Specify video format to use (default is 299 (1080P))
                        Other common formats are: 401 (2160P), 400 (1440P), 136 (720P), 397 (480P), 396 (360P)
                        Note selecting a format that is not provided will result in the script crashing non gracefully
  -v, --verbose
  -s, --short           Will produce a short pptx just containing the last frame of each specified clip
  -r, --reverse         Development option: Reverses slide order for faster iteration
  --no-output           Development option: Skips saving the final pptx to the hard drive.
```
## License Notice
Do not you this tool to infringe Copyright of other entities!
Only convert videos where you have the explicit right to use the content or use your own content.

The example file uses one my personal videos, you may also convert the given video into a presentation.

