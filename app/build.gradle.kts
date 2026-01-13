plugins {
    id("com.android.application")
    id("org.jetbrains.kotlin.android")
}

repositories {
    google()
    mavenCentral()
}

android {

    repositories {
        google()
        mavenCentral()
    }
    namespace = "com.example.qrventory"
    compileSdk = 34

    defaultConfig {
        applicationId = "com.example.qrventory"
        minSdk = 26
        targetSdk = 34
        versionCode = 1
        versionName = "1.0"

        multiDexEnabled = true
    }

    packaging {
        resources {
            excludes += setOf(
                "META-INF/DEPENDENCIES",
                "META-INF/DEPENDENCIES.txt",
                "META-INF/LICENSE",
                "META-INF/LICENSE.txt",
                "META-INF/NOTICE",
                "META-INF/NOTICE.txt"
            )
        }
    }

    buildTypes {
        getByName("release") {
            isMinifyEnabled = false
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro"
            )
        }
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }

    kotlinOptions {
        jvmTarget = "17"
    }

    buildFeatures {
        viewBinding = true
    }
}

configurations.all {
    resolutionStrategy {
        force("com.google.api-client:google-api-client:1.33.0")
    }
}


dependencies {
    // Core de Android y UI
    implementation("androidx.core:core-ktx:1.12.0")
    implementation("androidx.appcompat:appcompat:1.6.1")
    implementation("com.google.android.material:material:1.11.0")
    implementation("androidx.constraintlayout:constraintlayout:2.1.4")
    implementation("androidx.multidex:multidex:2.0.1")

    // ML Kit
    implementation("com.google.mlkit:barcode-scanning:17.2.0")

    // CameraX
    implementation("androidx.camera:camera-core:1.2.3")
    implementation("androidx.camera:camera-camera2:1.2.3")
    implementation("androidx.camera:camera-lifecycle:1.2.3")
    implementation("androidx.camera:camera-view:1.2.3")

    // Google Sign-In
    implementation("com.google.android.gms:play-services-auth:21.0.0")

    // Apache POI para archivos Excel locales
    implementation("org.apache.poi:poi-ooxml:5.2.3")

    // HTTP client y JSON para Google APIs
    implementation("com.google.api-client:google-api-client-android:1.33.0")
    implementation("com.google.http-client:google-http-client-android:1.43.3")
    implementation("com.google.http-client:google-http-client-gson:1.40.1")

    // OAuth en Android
    implementation("com.google.oauth-client:google-oauth-client:1.33.0")
    implementation("com.google.api-client:google-api-client:1.33.0")

    // Google Drive API (Java client)
    implementation("com.google.apis:google-api-services-drive:v3-rev20230815-2.0.0")

    // Google Sheets API (versión rev. más reciente en mavenCentral)
    implementation("com.google.apis:google-api-services-sheets:v4-rev20210629-1.32.1")

    // Gson para JSON
    implementation("com.google.code.gson:gson:2.10.1")

    // Kotlin coroutines
    implementation("org.jetbrains.kotlinx:kotlinx-coroutines-android:1.7.3")

    testImplementation("junit:junit:4.13.2")
}

