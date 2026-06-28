import CommonFormats from "src/CommonFormats.ts";
import type { FileData, FileFormat, FormatHandler } from "../FormatHandler.ts";
import { SimpleTTS } from "./espeakng.js/js/espeakng-simple.js";
import { WaveFile } from "wavefile";
import { appAssetUrl } from "../assetUrl.ts";

export class espeakngHandler implements FormatHandler {
  public name: string = "espeakng";
  public ready: boolean = true;
  #tts: SimpleTTS | undefined = undefined;

  public supportedFormats: FileFormat[] = [
    CommonFormats.TEXT.supported("text", true, false),
    CommonFormats.WAV.supported("wav", false, true)
  ];

  async init() {
    this.ready = true;
  }

  // Lazy-load the TTS worker so it does not slow down initial format listing.
  async getTTS(): Promise<SimpleTTS> {
    if(this.#tts == undefined) {
      await new Promise<void>((resolve, reject) => {
        this.#tts = new SimpleTTS({
          workerPath: appAssetUrl("js/espeakng.worker.js"),
          defaultVoice: "en",
          defaultRate: 220,
          defaultPitch: 200,
          enhanceAudio: true
        });
        this.#tts.onReady((error?: unknown) => {
          if (error) reject(error);
          else resolve();
        })
      });
    }
    return this.#tts!;
  }

  async doConvert (
    inputFiles: FileData[],
    inputFormat: FileFormat,
    outputFormat: FileFormat
  ): Promise<FileData[]> {
    if (inputFormat.mime !== CommonFormats.TEXT.mime || outputFormat.mime !== CommonFormats.WAV.mime)
      throw new Error("Unsupported conversion.");

    const tts = await this.getTTS();
    return Promise.all(inputFiles.map(async(file) => {
      const audio = await new Promise<AudioBuffer>((resolve, reject) => {
        tts.speak(new TextDecoder().decode(file.bytes), (audio: Float32Array | null) => {
          if (!audio) reject(new Error("Text to speech returned no audio."));
          else resolve(SimpleTTS.createAudioBuffer(audio, tts.sampleRate) as AudioBuffer);
        })
      });
      const samples = audio.getChannelData(0);
      const wav = new WaveFile();
      // Increasing pitch doesn't seem to do anything, so instead we
      // decrease playback rate and increase playback sample rate.
      wav.fromScratch(1, tts.sampleRate * 1.4, "32f", samples);
      return {
        name: file.name.split(".").slice(0, -1).join(".")+".wav",
        bytes: new Uint8Array(wav.toBuffer())
      }
    }))
  }
}

export default espeakngHandler;
