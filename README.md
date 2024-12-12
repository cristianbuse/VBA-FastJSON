# VBA-FastJSON
Fast Native JSON Parser / Serializer for VBA. Compatible with Windows and Mac.

## Installation

Download the latest [release](https://github.com/cristianbuse/VBA-FastJSON/releases), extract and import the ```LibJSON.bas``` module into your project.

Additionally, a ```Dictionary``` is required. While you can use Scripting.Dictionary (Microsoft Scripting Runtime reference - scrrun.dll on Windows), it is recommended to use [VBA-FastDictionary](https://github.com/cristianbuse/VBA-FastDictionary) because:
- it is Mac compatible
- faster in almost every way - see [Benchmarking VBA-FastDictionary](https://github.com/cristianbuse/VBA-FastDictionary/blob/master/benchmarking/README.md)
- hashes numbers better
For more information see [Cons of Scripting.Dictionary](https://github.com/cristianbuse/VBA-FastDictionary/blob/master/benchmarking/README.md#scriptingdictionary)

## Implementation

Non-Recursive JSON Parser.