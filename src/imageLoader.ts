// src/imageLoader.ts
// 이미지 로딩 및 변환 유틸리티

import * as fs from "fs";
import * as path from "path";

/**
 * 이미지 로드 결과
 */
export interface ImageLoadResult {
  buffer: Buffer;
  mimeType: string; // "image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp"
  success: boolean;
  error?: string;
}

/**
 * 파일 경로에서 MIME 타입 추론
 */
function getMimeType(filePath: string): string {
  const ext = path.extname(filePath).toLowerCase();
  const mimeTypes: Record<string, string> = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".webp": "image/webp",
  };
  return mimeTypes[ext] || "image/png";
}

/**
 * URL이 HTTP/HTTPS인지 확인
 */
function isHttpUrl(src: string): boolean {
  try {
    const url = new URL(src);
    return url.protocol === "http:" || url.protocol === "https:";
  } catch {
    return false;
  }
}

/**
 * 로컬 파일에서 이미지 로드
 */
export function loadLocalImage(filePath: string, basePath?: string): ImageLoadResult {
  try {
    // 상대 경로인 경우 basePath와 결합
    let resolvedPath = filePath;
    if (basePath && !path.isAbsolute(filePath)) {
      resolvedPath = path.join(basePath, filePath);
    }

    // 파일 존재 확인
    if (!fs.existsSync(resolvedPath)) {
      return {
        buffer: Buffer.alloc(0),
        mimeType: "image/png",
        success: false,
        error: `File not found: ${resolvedPath}`,
      };
    }

    // 파일 읽기
    const buffer = fs.readFileSync(resolvedPath);
    const mimeType = getMimeType(resolvedPath);

    return {
      buffer,
      mimeType,
      success: true,
    };
  } catch (err) {
    const errorMsg = err instanceof Error ? err.message : "Unknown error";
    return {
      buffer: Buffer.alloc(0),
      mimeType: "image/png",
      success: false,
      error: `Failed to load image: ${errorMsg}`,
    };
  }
}

/**
 * HTTP(S) URL에서 이미지 다운로드 (Node.js built-in fetch 또는 https/http 모듈)
 * Node.js 18.x+ 사용 시 fetch 지원
 */
export async function loadRemoteImage(url: string): Promise<ImageLoadResult> {
  try {
    // Node.js 18+ fetch 사용
    if (typeof fetch !== "undefined") {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 10000);

      try {
        const response = await fetch(url, {
          signal: controller.signal,
        });

        clearTimeout(timeoutId);

        if (!response.ok) {
          return {
            buffer: Buffer.alloc(0),
            mimeType: "image/png",
            success: false,
            error: `HTTP ${response.status}: ${response.statusText}`,
          };
        }

        const arrayBuffer = await response.arrayBuffer();
        const buffer = Buffer.from(arrayBuffer);
        const contentType = response.headers.get("content-type") || "image/png";
        const mimeType = contentType.split(";")[0].trim();

        return {
          buffer,
          mimeType,
          success: true,
        };
      } catch (fetchErr) {
        clearTimeout(timeoutId);
        throw fetchErr;
      }
    }

    // 폴백: https/http 모듈 사용 (Node.js < 18)
    return await loadRemoteImageFallback(url);
  } catch (err) {
    const errorMsg = err instanceof Error ? err.message : "Unknown error";
    return {
      buffer: Buffer.alloc(0),
      mimeType: "image/png",
      success: false,
      error: `Failed to download image: ${errorMsg}`,
    };
  }
}

// CommonJS require 모듈 (Node.js http/https 서버 요청용)
// import 대신 require 사용: Node.js 내장 모듈 CommonJS 호환성
const https = require("https");
const http = require("http");

/**
 * https/http 모듈을 사용한 폴백 다운로드
 */
async function loadRemoteImageFallback(urlString: string): Promise<ImageLoadResult> {
  return new Promise((resolve) => {
    const url = new URL(urlString);
    const isHttps = url.protocol === "https:";
    const httpModule = isHttps ? https : http;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const request = httpModule.get(urlString, { timeout: 10000 }, (response: any) => {
      const chunks: Buffer[] = [];

      if (response.statusCode !== 200) {
        resolve({
          buffer: Buffer.alloc(0),
          mimeType: "image/png",
          success: false,
          error: `HTTP ${response.statusCode}`,
        });
        response.resume(); // consume response to free up memory
        return;
      }

      response.on("data", (chunk: Buffer) => {
        chunks.push(chunk);
      });

      response.on("end", () => {
        const buffer = Buffer.concat(chunks);
        const contentType = response.headers["content-type"] || "image/png";
        const mimeType = contentType.split(";")[0].trim();

        resolve({
          buffer,
          mimeType,
          success: true,
        });
      });

      response.on("error", (err: Error) => {
        resolve({
          buffer: Buffer.alloc(0),
          mimeType: "image/png",
          success: false,
          error: `Download error: ${err.message}`,
        });
      });
    });

    request.on("error", (err: Error) => {
      resolve({
        buffer: Buffer.alloc(0),
        mimeType: "image/png",
        success: false,
        error: `Request error: ${err.message}`,
      });
    });

    request.on("timeout", () => {
      request.destroy();
      resolve({
        buffer: Buffer.alloc(0),
        mimeType: "image/png",
        success: false,
        error: "Request timeout (10s)",
      });
    });
  });
}

/**
 * 이미지 소스(경로 또는 URL)에서 이미지 로드 (동기/비동기 래퍼)
 */
export async function loadImage(src: string, basePath?: string): Promise<ImageLoadResult> {
  if (isHttpUrl(src)) {
    return loadRemoteImage(src);
  }
  // 로컬 파일은 동기 로드
  return loadLocalImage(src, basePath);
}

/**
 * MIME 타입에서 docx Media 확장자 추론
 */
export function getDocxMediaExtension(mimeType: string): string {
  const extensions: Record<string, string> = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/webp": "webp",
  };
  return extensions[mimeType] || "png";
}
