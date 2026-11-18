// AppKit checks without synthetic mouse events or visible frame changes.
#import <AppKit/AppKit.h>
#include <node_api.h>

@interface NSObject (GenOfficeGestureTest)
- (instancetype)initWithWindow:(NSWindow*)window;
- (void)finish;
- (void)invalidate;
@end

static napi_value Run(napi_env env, napi_callback_info info) {
  size_t count = 1, length = 0;
  napi_value args[1]; void* bytes = nullptr; bool buffer = false;
  napi_get_cb_info(env, info, &count, args, nullptr, nullptr);
  if (count != 1 || napi_is_buffer(env, args[0], &buffer) != napi_ok || !buffer ||
      napi_get_buffer_info(env, args[0], &bytes, &length) != napi_ok || length != sizeof(void*) ||
      !NSThread.isMainThread) {
    napi_throw_error(env, nullptr, "Expected a native window on the main thread");
    return nullptr;
  }
  NSView* view = (__bridge NSView*)*(void**)bytes;
  NSWindow* window = view.window;
  Class type = NSClassFromString(@"GenOfficeGestureGuard");
  if (!type || !window) {
    napi_throw_error(env, nullptr, "Production gesture guard was not loaded"); return nullptr;
  }
  id guard = [[type alloc] initWithWindow:window];
  NSRect original = window.frame;
  NSMutableDictionary* checks = [NSMutableDictionary dictionary];
  NSTimer* timer = [NSTimer timerWithTimeInterval:60 repeats:NO block:^(NSTimer*) {}];
  [guard setValue:timer forKey:@"timer"];
  [NSNotificationCenter.defaultCenter postNotificationName:NSWindowWillMoveNotification object:window];
  checks[@"nativeStartCancelsBeforeFrameChange"] = @([guard valueForKey:@"timer"] == nil);
  checks[@"nativeStartInvalidatesTimer"] = @(!timer.valid);
  checks[@"nativeStartRecorded"] = @([[guard valueForKey:@"nativeWins"] unsignedIntegerValue] == 1);
  checks[@"nativeStartStatus"] = @([[guard valueForKey:@"status"] isEqual:@"native-started"]);
  checks[@"nativeStartLeavesFrameUntouched"] = @(NSEqualRects(window.frame, original));
  timer = [NSTimer timerWithTimeInterval:60 repeats:NO block:^(NSTimer*) {}];
  [guard setValue:timer forKey:@"timer"];
  [guard setValue:@YES forKey:@"rescuing"];
  [NSNotificationCenter.defaultCenter postNotificationName:NSWindowWillMoveNotification object:window];
  checks[@"recoveryDoesNotCancelItself"] = @([guard valueForKey:@"timer"] == timer && timer.valid);
  checks[@"recoveryNotCountedAsNative"] = @([[guard valueForKey:@"nativeWins"] unsignedIntegerValue] == 1);
  [guard setValue:@NO forKey:@"rescuing"];
  [guard setValue:@120 forKey:@"maxDistance"];
  [guard setValue:@42 forKey:@"lastEventNumber"];
  [guard finish];
  checks[@"releaseStopsTimer"] = @([guard valueForKey:@"timer"] == nil && !timer.valid);
  checks[@"releasePreservesMissedGesture"] = @([[guard valueForKey:@"missedDrags"] unsignedIntegerValue] == 1 && [[guard valueForKey:@"lastMissedEventNumber"] integerValue] == 42);
  [guard invalidate];
  checks[@"cleanupRemovesObservers"] = @([guard valueForKey:@"monitor"] == nil && [guard valueForKey:@"moveObserver"] == nil && [guard valueForKey:@"closeObserver"] == nil);
  checks[@"testLeavesFrameUntouched"] = @(NSEqualRects(window.frame, original));
  NSData* data = [NSJSONSerialization dataWithJSONObject:checks options:0 error:nil];
  napi_value result;
  napi_create_string_utf8(env, (const char*)data.bytes, data.length, &result);
  return result;
}
static napi_value Init(napi_env env, napi_value exports) {
  napi_value fn; napi_create_function(env, "run", NAPI_AUTO_LENGTH, Run, nullptr, &fn);
  napi_set_named_property(env, exports, "run", fn); return exports;
}
NAPI_MODULE(NODE_GYP_MODULE_NAME, Init)
