#!/bin/sh
# CI uses the Gradle Actions setup and invokes `gradle` directly.
# This file is kept for Android Studio/GitHub project structure compatibility.
exec gradle "$@"
