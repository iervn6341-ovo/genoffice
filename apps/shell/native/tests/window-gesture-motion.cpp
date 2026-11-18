#include "../window-gesture-motion.h"
#include <cassert>
#include <cstdio>

int main() {
  GenOfficeDragMotion motion;
  // A long stationary hold does not consume the native gesture's grace period.
  assert(!motion.ready(0, 0));
  assert(!motion.ready(10, 0));
  assert(!motion.ready(11, 20));
  assert(!motion.ready(11.04, 40));
  assert(motion.ready(11.06, 50));
  // Returning to the press point resets the stall, rather than arming a future drag.
  assert(!motion.ready(12, 2));
  assert(!motion.ready(14, 20));
  assert(!motion.ready(14.04, 30));
  assert(motion.ready(14.06, 40));
  // Tiny motion and a canceled/replaced gesture never inherit a previous timer.
  motion = {};
  assert(!motion.ready(20, 5.9));
  assert(!motion.ready(30, 5.9));
  assert(!motion.ready(31, 6));
  assert(!motion.ready(31.04, 8));
  assert(motion.ready(31.06, 12));
  // R3 physical event 15072: release at 237 ms. Motion should recover while
  // still held, rather than waiting 200 ms from its first meaningful motion.
  motion = {};
  assert(!motion.ready(0, 0));
  assert(!motion.ready(0.05, 8));
  assert(motion.ready(0.117, 120));
  // Native will-move wins even when AppKit has not updated the frame yet.
  motion = {};
  assert(!motion.ready(0, 8));
  motion.nativeStarted = true;
  assert(!motion.ready(0.06, 60));
  assert(!motion.ready(1, 100));
  // Only a new eligible press resets the native gesture's ownership.
  motion = {};
  assert(!motion.ready(2, 10));
  assert(motion.ready(2.06, 40));
  std::puts("Native drag grace-period checks passed");
}
