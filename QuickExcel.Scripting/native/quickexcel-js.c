#include <windows.h>
#include <stdint.h>
#include <stdlib.h>
#include <string.h>
#include "quickjs.h"

#define QX_EXPORT __declspec(dllexport)
#define QX_RESULT_LIMIT (4 * 1024 * 1024)

typedef struct {
    JSRuntime *runtime;
    JSContext *context;
    JSValue bridge;
    ULONGLONG deadline;
    int interrupted;
} qx_session;

static char *copy_text(const char *text, size_t length) {
    char *copy = malloc(length + 1);
    if (copy) { memcpy(copy, text, length); copy[length] = 0; }
    return copy;
}

static void set_error(char **error, const char *message) {
    if (error) *error = copy_text(message, strlen(message));
}

static int interrupt(JSRuntime *runtime, void *opaque) {
    (void)runtime;
    qx_session *session = opaque;
    if (GetTickCount64() >= session->deadline) {
        session->interrupted = 1;
        return 1;
    }
    return 0;
}

static void start_operation(qx_session *session) {
    JS_UpdateStackTop(session->runtime);
    session->interrupted = 0;
    session->deadline = GetTickCount64() + 1000;
}

static void exception_error(qx_session *session, char **error) {
    JSValue exception = JS_GetException(session->context);
    if (session->interrupted) {
        set_error(error, "JavaScript execution exceeded the 1 second time limit.");
    } else {
        size_t length = 0;
        const char *text = JS_ToCStringLen(session->context, &length, exception);
        if (text) {
            if (length > 16384) length = 16384;
            if (error) *error = copy_text(text, length);
            JS_FreeCString(session->context, text);
        } else {
            set_error(error, "JavaScript failed (execution or memory limit reached).");
            JSValue secondary = JS_GetException(session->context);
            JS_FreeValue(session->context, secondary);
        }
    }
    JS_FreeValue(session->context, exception);
}

QX_EXPORT void qx_destroy(qx_session *session) {
    if (!session) return;
    if (session->context) {
        JS_FreeValue(session->context, session->bridge);
        JS_FreeContext(session->context);
    }
    if (session->runtime) JS_FreeRuntime(session->runtime);
    free(session);
}

QX_EXPORT qx_session *qx_create(const char *bootstrap, char **error) {
    if (error) *error = NULL;
    qx_session *session = calloc(1, sizeof(qx_session));
    if (!session) { set_error(error, "Unable to allocate JavaScript session."); return NULL; }
    session->bridge = JS_UNDEFINED;
    session->runtime = JS_NewRuntime();
    if (!session->runtime) { set_error(error, "Unable to create QuickJS runtime."); qx_destroy(session); return NULL; }
    JS_SetMemoryLimit(session->runtime, 64 * 1024 * 1024);
    JS_SetMaxStackSize(session->runtime, 256 * 1024);
    JS_SetInterruptHandler(session->runtime, interrupt, session);
    session->context = JS_NewContext(session->runtime);
    if (!session->context) { set_error(error, "Unable to create QuickJS context."); qx_destroy(session); return NULL; }
    start_operation(session);
    session->bridge = JS_Eval(session->context, bootstrap, strlen(bootstrap), "quickexcel-host.js", JS_EVAL_TYPE_GLOBAL);
    if (JS_IsException(session->bridge)) {
        exception_error(session, error);
        session->bridge = JS_UNDEFINED;
        qx_destroy(session);
        return NULL;
    }
    return session;
}

QX_EXPORT char *qx_call(qx_session *session, const char *method, const char *arguments_json, char **error) {
    if (error) *error = NULL;
    if (!session) { set_error(error, "JavaScript session is closed."); return NULL; }
    start_operation(session);
    JSContext *context = session->context;
    JSValue arguments = JS_ParseJSON(context, arguments_json, strlen(arguments_json), "arguments.json");
    if (JS_IsException(arguments)) { exception_error(session, error); return NULL; }
    JSValue function = JS_GetPropertyStr(context, session->bridge, method);
    JSValue argv[2] = { JS_GetPropertyUint32(context, arguments, 0), JS_GetPropertyUint32(context, arguments, 1) };
    JSValue result = JS_Call(context, function, session->bridge, 2, argv);
    JS_FreeValue(context, argv[0]);
    JS_FreeValue(context, argv[1]);
    JS_FreeValue(context, function);
    JS_FreeValue(context, arguments);
    if (JS_IsException(result)) { exception_error(session, error); return NULL; }
    JSValue json = JS_JSONStringify(context, result, JS_UNDEFINED, JS_UNDEFINED);
    JS_FreeValue(context, result);
    if (JS_IsException(json)) { exception_error(session, error); return NULL; }
    size_t length = 0;
    const char *text = JS_ToCStringLen(context, &length, json);
    char *output = NULL;
    if (text) {
        if (length <= QX_RESULT_LIMIT) {
            output = copy_text(text, length);
            if (!output) set_error(error, "Unable to allocate JavaScript result.");
        } else {
            set_error(error, "JavaScript result exceeds the 4 MiB result limit.");
        }
        JS_FreeCString(context, text);
    } else {
        exception_error(session, error);
    }
    JS_FreeValue(context, json);
    return output;
}

QX_EXPORT void qx_free_string(char *text) { free(text); }
