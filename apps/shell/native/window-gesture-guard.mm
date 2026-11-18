// App-local fallback for stalled AppKit title-bar gestures on macOS 27.2.
// Events are observed, never suppressed or reposted. Native movement wins.
#import <AppKit/AppKit.h>
#import <objc/runtime.h>
#include <node_api.h>
#include <algorithm>
#include <cmath>
#include "window-gesture-motion.h"

static char kGuardKey;

@interface GenOfficeGestureGuard : NSObject {
  GenOfficeDragMotion motion_;
}
@property(nonatomic, weak) NSWindow* window;
@property(nonatomic, strong) id monitor;
@property(nonatomic, strong) id closeObserver;
@property(nonatomic, strong) id moveObserver;
@property(nonatomic, strong) NSTimer* timer;
@property(nonatomic) NSPoint startMouse;
@property(nonatomic) NSRect startFrame;
@property(nonatomic) BOOL rescuing;
@property(nonatomic) NSUInteger attempts;
@property(nonatomic) NSUInteger rescues;
@property(nonatomic) NSUInteger presses;
@property(nonatomic) NSUInteger ticks;
@property(nonatomic) NSUInteger nativeWins;
@property(nonatomic) NSUInteger missedDrags;
@property(nonatomic) NSInteger lastMissedEventNumber;
@property(nonatomic) NSInteger lastEventNumber;
@property(nonatomic) CGFloat maxDistance;
@property(nonatomic, copy) NSString* status;
- (instancetype)initWithWindow:(NSWindow*)window;
- (void)stop;
- (void)invalidate;
- (void)begin:(NSEvent*)event;
- (void)tick;
- (void)nativeStarted;
- (void)finish;
@end

@implementation GenOfficeGestureGuard
- (instancetype)initWithWindow:(NSWindow*)window {
  if ((self = [super init])) {
    self.window = window;
    self.status = @"idle";
    __weak GenOfficeGestureGuard* weakSelf = self;
    self.monitor = [NSEvent addLocalMonitorForEventsMatchingMask:
        NSEventMaskLeftMouseDown | NSEventMaskLeftMouseUp
        handler:^NSEvent*(NSEvent* event) {
      GenOfficeGestureGuard* guard = weakSelf;
      if (!guard || event.window != guard.window) return event;
      if (event.type == NSEventTypeLeftMouseUp) [guard finish];
      else [guard begin:event];
      return event;
    }];
    self.closeObserver = [NSNotificationCenter.defaultCenter
        addObserverForName:NSWindowWillCloseNotification object:window queue:nil
        usingBlock:^(NSNotification* note) { [weakSelf invalidate]; }];
    self.moveObserver = [NSNotificationCenter.defaultCenter
        addObserverForName:NSWindowWillMoveNotification object:window queue:nil
        usingBlock:^(NSNotification* note) { [weakSelf nativeStarted]; }];
  }
  return self;
}

- (void)begin:(NSEvent*)event {
  [self stop];
  self.presses++;
  self.lastEventNumber = event.eventNumber;
  self.maxDistance = 0;
  self.status = @"ineligible";
  NSWindow* window = self.window;
  if (!window || !window.movable ||
      (window.styleMask & NSWindowStyleMaskFullScreen) ||
      !(NSEvent.pressedMouseButtons & 1) ||
      (event.modifierFlags & NSEventModifierFlagControl) || event.clickCount != 1) return;
  const NSRect frame = window.frame;
  const NSPoint point = event.locationInWindow;
  const CGFloat fromTop = NSHeight(frame) - point.y;
  // Keep traffic lights, tabs and every native resize edge out of this path.
  if (point.x < 6 || point.x > NSWidth(frame) - 6 || fromTop < 6 || fromTop > 40) return;
  NSView* frameView = window.contentView.superview;
  if (!frameView || [frameView hitTest:point] != frameView) return;
  self.attempts++;
  self.status = @"native-grace";
  self.startMouse = NSEvent.mouseLocation;
  self.startFrame = frame;
  motion_ = {};
  __weak GenOfficeGestureGuard* weakSelf = self;
  self.timer = [NSTimer timerWithTimeInterval:1.0 / 60.0 repeats:YES block:^(NSTimer* timer) {
    [weakSelf tick];
  }];
  // Only a held title-bar gesture has a timer; nothing runs while idle.
  [NSRunLoop.mainRunLoop addTimer:self.timer forMode:NSRunLoopCommonModes];
}

- (void)tick {
  self.ticks++;
  NSWindow* window = self.window;
  if (!window || !window.isVisible || !window.isKeyWindow || !window.movable ||
      (window.styleMask & NSWindowStyleMaskFullScreen)) {
    self.status = @"gesture-ended";
    [self stop];
    return;
  }
  if (!(NSEvent.pressedMouseButtons & 1)) {
    [self finish];
    return;
  }
  // Any native geometry update cancels the fallback for this entire gesture.
  if (!self.rescuing && !NSEqualRects(window.frame, self.startFrame)) {
    [self nativeStarted];
    return;
  }
  NSPoint mouse = NSEvent.mouseLocation;
  const CGFloat dx = mouse.x - self.startMouse.x;
  const CGFloat dy = mouse.y - self.startMouse.y;
  self.maxDistance = std::max(self.maxDistance, std::hypot(dx, dy));
  if (!self.rescuing) {
    // Count the grace period from motion, not mouse-down: a long stationary
    // press must still give the native gesture time to begin once it moves.
    const NSTimeInterval now = NSProcessInfo.processInfo.systemUptime;
    if (!motion_.ready(now, std::hypot(dx, dy))) return;
    self.rescuing = YES;
    self.rescues++;
    self.status = @"recovered";
  }
  NSPoint origin = NSMakePoint(std::round(self.startFrame.origin.x + dx),
                              std::round(self.startFrame.origin.y + dy));
  // Preserve cross-display coordinates but keep the title bar below the menu.
  for (NSScreen* screen in NSScreen.screens) {
    if (NSPointInRect(mouse, screen.frame)) {
      origin.y = std::min(origin.y, NSMaxY(screen.visibleFrame) - NSHeight(window.frame));
      break;
    }
  }
  if (!NSEqualPoints(window.frame.origin, origin)) [window setFrameOrigin:origin];
}

- (void)nativeStarted {
  // AppKit can announce a drag before the first actual move. Yield now, rather
  // than racing that first frame with the shorter stall deadline. Our own
  // setFrameOrigin calls must not cancel a recovery already in progress.
  if (!self.timer || self.rescuing) return;
  motion_.nativeStarted = true;
  self.nativeWins++;
  self.status = @"native-started";
  [self stop];
}

- (void)finish {
  if (self.timer && !self.rescuing && self.maxDistance >= 6) {
    self.missedDrags++;
    self.lastMissedEventNumber = self.lastEventNumber;
    self.status = @"released-before-recovery";
  }
  [self stop];
}

- (void)stop {
  [self.timer invalidate];
  self.timer = nil;
  self.rescuing = NO;
}

- (void)invalidate {
  [self stop];
  if (self.monitor) [NSEvent removeMonitor:self.monitor];
  if (self.closeObserver) [NSNotificationCenter.defaultCenter removeObserver:self.closeObserver];
  if (self.moveObserver) [NSNotificationCenter.defaultCenter removeObserver:self.moveObserver];
  self.monitor = nil;
  self.closeObserver = nil;
  self.moveObserver = nil;
}
@end

static NSWindow* WindowArgument(napi_env env, napi_callback_info info) {
  size_t argc = 1;
  napi_value args[1];
  void* data = nullptr;
  size_t length = 0;
  bool buffer = false;
  napi_get_cb_info(env, info, &argc, args, nullptr, nullptr);
  if (argc != 1 || napi_is_buffer(env, args[0], &buffer) != napi_ok || !buffer ||
      napi_get_buffer_info(env, args[0], &data, &length) != napi_ok || length != sizeof(void*)) {
    napi_throw_type_error(env, nullptr, "Expected an Electron native window handle");
    return nil;
  }
  if (!NSThread.isMainThread) return nil;
  NSView* view = (__bridge NSView*)*(void**)data;
  return view.window;
}

static napi_value Attach(napi_env env, napi_callback_info info) {
  NSWindow* window = WindowArgument(env, info);
  const NSOperatingSystemVersion os = NSProcessInfo.processInfo.operatingSystemVersion;
  const bool enabled = window && os.majorVersion == 27 && os.minorVersion == 2;
  if (enabled && !objc_getAssociatedObject(window, &kGuardKey)) {
    GenOfficeGestureGuard* guard = [[GenOfficeGestureGuard alloc] initWithWindow:window];
    objc_setAssociatedObject(window, &kGuardKey, guard, OBJC_ASSOCIATION_RETAIN_NONATOMIC);
  }
  napi_value result;
  napi_get_boolean(env, enabled, &result);
  return result;
}

static napi_value Stats(napi_env env, napi_callback_info info) {
  NSWindow* window = WindowArgument(env, info);
  GenOfficeGestureGuard* guard = window ? objc_getAssociatedObject(window, &kGuardKey) : nil;
  napi_value result, value;
  napi_create_object(env, &result);
  napi_get_boolean(env, guard != nil, &value);
  napi_set_named_property(env, result, "installed", value);
  napi_create_uint32(env, (uint32_t)guard.attempts, &value);
  napi_set_named_property(env, result, "attempts", value);
  napi_create_uint32(env, (uint32_t)guard.rescues, &value);
  napi_set_named_property(env, result, "rescues", value);
  // Bounded diagnostic counters only: no document data, event log or idle work.
  napi_create_uint32(env, (uint32_t)guard.presses, &value);
  napi_set_named_property(env, result, "presses", value);
  napi_create_uint32(env, (uint32_t)guard.ticks, &value);
  napi_set_named_property(env, result, "ticks", value);
  napi_create_uint32(env, (uint32_t)guard.nativeWins, &value);
  napi_set_named_property(env, result, "nativeWins", value);
  napi_create_uint32(env, (uint32_t)guard.missedDrags, &value);
  napi_set_named_property(env, result, "missedDrags", value);
  napi_create_int64(env, guard.lastMissedEventNumber, &value);
  napi_set_named_property(env, result, "lastMissedEventNumber", value);
  napi_create_int64(env, guard.lastEventNumber, &value);
  napi_set_named_property(env, result, "lastEventNumber", value);
  napi_create_double(env, guard.maxDistance, &value);
  napi_set_named_property(env, result, "maxDistance", value);
  napi_create_string_utf8(env, guard.status.UTF8String ?: "uninstalled", NAPI_AUTO_LENGTH, &value);
  napi_set_named_property(env, result, "status", value);
  napi_get_boolean(env, guard.timer != nil, &value);
  napi_set_named_property(env, result, "active", value);
  return result;
}

static napi_value Init(napi_env env, napi_value exports) {
  napi_value fn;
  napi_create_function(env, "attach", NAPI_AUTO_LENGTH, Attach, nullptr, &fn);
  napi_set_named_property(env, exports, "attach", fn);
  napi_create_function(env, "stats", NAPI_AUTO_LENGTH, Stats, nullptr, &fn);
  napi_set_named_property(env, exports, "stats", fn);
  return exports;
}
NAPI_MODULE(NODE_GYP_MODULE_NAME, Init)
