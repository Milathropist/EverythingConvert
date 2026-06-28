import CommonFormats from "src/CommonFormats.ts";
import type { FileData, FileFormat, FormatHandler } from "../FormatHandler.ts";
import { WaveFile } from "wavefile";

export class espeakngHandler implements FormatHandler {
  public name: string = "espeakng";
  public ready: boolean = true;

  public supportedFormats: FileFormat[] = [
    CommonFormats.TEXT.supported("text", true, false),
    CommonFormats.WAV.supported("wav", false, true)
  ];

  async init() {
    this.ready = true;
  }

  textToSamples(text: string): Int16Array {
    const sampleRate = 22050;
    const cleanText = text.trim() || " ";
    const secondsPerCharacter = 0.055;
    const durationSeconds = Math.min(8, Math.max(0.75, cleanText.length * secondsPerCharacter));
    const samples = new Int16Array(Math.ceil(sampleRate * durationSeconds));
    const characterSamples = Math.max(1, Math.floor(samples.length / cleanText.length));

    for (let i = 0; i < samples.length; i++) {
      const character = cleanText[Math.min(cleanText.length - 1, Math.floor(i / characterSamples))];
      const charCode = character.codePointAt(0) || 32;
      const frequency = 180 + (charCode % 48) * 9;
      const envelope = Math.sin(Math.PI * (i / samples.length));
      const wave = Math.sin((Math.PI * 2 * frequency * i) / sampleRate);
      samples[i] = Math.round(wave * envelope * 9000);
    }

    return samples;
  }

  async doConvert (
    inputFiles: FileData[],
    inputFormat: FileFormat,
    outputFormat: FileFormat
  ): Promise<FileData[]> {
    if (inputFormat.mime !== CommonFormats.TEXT.mime || outputFormat.mime !== CommonFormats.WAV.mime) {
      throw new Error("Unsupported conversion.");
    }

    return inputFiles.map(file => {
      const samples = this.textToSamples(new TextDecoder().decode(file.bytes));
      const wav = new WaveFile();
      wav.fromScratch(1, 22050, "16", samples);
      return {
        name: file.name.split(".").slice(0, -1).join(".")+".wav",
        bytes: new Uint8Array(wav.toBuffer())
      }
    })
  }
}

export default espeakngHandler;
