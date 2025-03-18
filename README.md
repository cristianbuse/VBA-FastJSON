# VBA-FastJSON
Fast Native JSON Parser / Serializer for VBA. Compatible with Windows and Mac.

[RFC 8259](https://datatracker.ietf.org/doc/html/rfc8259) compliant.

## Installation

Download the [latest release](https://github.com/cristianbuse/VBA-FastJSON/releases/latest), extract and import the ```LibJSON.bas``` module into your project.

Additionally, a ```Dictionary``` is required. While you can use Scripting.Dictionary (Microsoft Scripting Runtime reference - scrrun.dll on Windows), it is recommended to use [VBA-FastDictionary](https://github.com/cristianbuse/VBA-FastDictionary) because:
- it is Mac compatible
- faster in almost every way - see [Benchmarking VBA-FastDictionary](https://github.com/cristianbuse/VBA-FastDictionary/blob/master/benchmarking/README.md)
- hashes numbers better
- allows endless nesting

For more information see [Cons of Scripting.Dictionary](https://github.com/cristianbuse/VBA-FastDictionary/blob/master/benchmarking/README.md#scriptingdictionary)

## Parser

Non-Recursive implementation. 

## Testing

Download the [latest release](https://github.com/cristianbuse/VBA-FastJSON/releases/latest), extract and import the ```TestLibJSON.bas``` module into your project. Run ```RunAllJSONTests``` method. On failure, execution will stop on the first failed ```Assert```.

Many thanks to Nicolas Seriot ([@nst](https://github.com/nst)). This repo includes some of the tests found at [JSONTestSuite](https://github.com/nst/JSONTestSuite).

Must-read his article: [Parsing JSON is a Minefield](https://seriot.ch/projects/parsing_json.html)!
