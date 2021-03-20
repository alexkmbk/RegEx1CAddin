#include "jnienv.h"

#include <stdlib.h>
#include <android/log.h>

static JavaVM* sJavaVM = NULL;

void trace(const char* format, ...)
{
    va_list args;
    va_start(args, format);
    __android_log_vprint(ANDROID_LOG_INFO, "RegEx", format, args);
    va_end(args);
}

JNIEnv* getJniEnv()
{
    trace("getJniEnv()"); 

    JNIEnv* env = NULL;

    switch (sJavaVM->GetEnv((void**)&env, JNI_VERSION_1_6))
    {
        case JNI_OK:
            trace("GetEnv(), env = %08X", env);
            return env;

        case JNI_EDETACHED:
        {
            JavaVMAttachArgs args;
            args.name = NULL;
            args.group = NULL;
            args.version = JNI_VERSION_1_6;

            if (!sJavaVM->AttachCurrentThreadAsDaemon(&env, &args))
            {
                trace("AttachCurrentThreadAsDaemon(), env = %08X", env);
                return env;
            }
            break;
        }

    }
    return NULL;
};

extern "C" JNIEXPORT jint JNICALL JNI_OnLoad(JavaVM* aJavaVM, void* aReserved)
{
    trace("JNI_OnLoad()");
    sJavaVM = aJavaVM;
    return JNI_VERSION_1_6;
}


extern "C" JNIEXPORT void JNICALL JNI_OnUnload(JavaVM* aJavaVM, void* aReserved)
{
	sJavaVM = NULL;
}

extern "C" JNIEXPORT void JNICALL Java_com_stepCounterPackage_stepCounterLib_StepCounterClass_NativeTrace(JNIEnv* jenv, jclass aClass, jstring aStr)
{
    const char* jchars = jenv->GetStringUTFChars(aStr, NULL); 
    trace(jchars);
    jenv->ReleaseStringUTFChars(aStr, jchars); 
}
