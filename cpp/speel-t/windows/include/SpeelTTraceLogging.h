#pragma once

#include <Windows.h>
#include <TraceLoggingActivity.h>
#include "sentry.h"

#ifdef __cplusplus
extern "C" {
#endif
TRACELOGGING_DECLARE_PROVIDER(g_hSpeelTLoggingProvider);
#ifdef __cplusplus
} // extern "C"
#endif