#pragma once

// AppKit's will-move notification wins before its first geometry update.
// A short grace after motion also covers quick drags that end within 200 ms.
struct GenOfficeDragMotion {
  double began = -1;
  bool nativeStarted = false;
  bool ready(double now, double distance) {
    if (nativeStarted) return false;
    if (distance < 6) { began = -1; return false; }
    if (began < 0) began = now;
    return now - began >= 0.05;
  }
};
