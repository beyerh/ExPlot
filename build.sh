#!/bin/bash
 
 # ExPlot build script (env creation + Nuitka build)
 
 # Configuration
 APP_NAME="ExPlot"
 VERSION="0.7.3"
 BUILD_DIR="build"
 
 VENV_APP=".venv"
 VENV_INTEL=".venv_x86"
 VENV_INTEL_COMPAT=".venv_x86_compat"
 
 # python.org universal2 Python path (recommended). Override by pasting a different python3 path.
 PYTHON_UNIVERSAL2="/Library/Frameworks/Python.framework/Versions/3.12/bin/python3"
 PYTHON_UNIVERSAL2_RELEASE_URL="https://www.python.org/downloads/release/python-31210/"
 PYTHON_UNIVERSAL2_RECOMMENDED_VERSION="3.12"
 
 # Intel build settings (for older macOS support)
 INTEL_COMPAT_REQUIREMENTS_FILE="requirements_intel_compatibility.txt"
 INTEL_DEPLOYMENT_TARGET="11.0"
 
 set -euo pipefail
 
 # Best-effort: if this script is launched from an active conda environment, try to
 # deactivate and/or neutralize it. Conda activation often causes non-portable
 # libpython link paths in produced binaries.
 if [[ -n "${CONDA_PREFIX:-}" || -n "${CONDA_DEFAULT_ENV:-}" ]]; then
   echo "Warning: Detected active conda environment ('${CONDA_DEFAULT_ENV:-unknown}'). Attempting to deactivate..." >&2
   if declare -F conda >/dev/null 2>&1; then
     conda deactivate >/dev/null 2>&1 || true
   fi
   unset CONDA_PREFIX CONDA_DEFAULT_ENV CONDA_PROMPT_MODIFIER CONDA_SHLVL CONDA_EXE CONDA_PYTHON_EXE
   unset _CE_CONDA _CE_M
 fi
 
 if [[ ! -t 0 ]]; then
   echo "Error: build.sh requires an interactive terminal (stdin is not a TTY)." >&2
   exit 2
 fi
 
 echo "ExPlot build wizard"
 echo
 
 TARGET=""
 while [[ -z "${TARGET}" ]]; do
   echo "Select target:"
   echo "  1) Apple Silicon (arm64)"
   echo "  2) Intel Mac (x86_64)"
   echo "  3) Intel compatibility (x86_64, macOS >= ${INTEL_DEPLOYMENT_TARGET})"
   read -r -p "Choice [1-3]: " target_choice
   case "${target_choice}" in
     1) TARGET="apple" ;;
     2) TARGET="intel" ;;
     3) TARGET="intel_compat" ;;
     *) echo "Invalid choice." ;;
   esac
 done

 MODE=""
 while [[ -z "${MODE}" ]]; do
   echo
   echo "Select action:"
   echo "  1) Recreate environment only (deletes existing venv)"
   echo "  2) Build app only"
   echo "  3) Recreate environment + build app (deletes existing venv)"
   read -r -p "Choice [1-3]: " mode_choice
   case "${mode_choice}" in
     1) MODE="env" ;;
     2) MODE="build" ;;
     3) MODE="all" ;;
     *) echo "Invalid choice." ;;
   esac
 done
 
 echo
 echo "Python used for creating venv: ${PYTHON_UNIVERSAL2}"
 read -r -p "Press Enter to accept, or paste a different python3 path: " python_override
 if [[ -n "${python_override}" ]]; then
   PYTHON_UNIVERSAL2="${python_override}"
 fi
 
 if [[ ! -x "${PYTHON_UNIVERSAL2}" ]]; then
   echo >&2
   echo "Error: Python executable not found at: ${PYTHON_UNIVERSAL2}" >&2
   echo "Recommended: install python.org 'universal2' Python ${PYTHON_UNIVERSAL2_RECOMMENDED_VERSION} (includes both arm64 and x86_64 support)." >&2
   echo "Download: ${PYTHON_UNIVERSAL2_RELEASE_URL}" >&2
   echo "More downloads: https://www.python.org/downloads/macos/" >&2
   echo >&2
   echo "What do you want to do?" >&2
   echo "  1) Install python.org universal2 first (recommended)" >&2
   echo "  2) Proceed using python3 from PATH" >&2
   echo "  3) Paste a custom python3 path" >&2
   read -r -p "Choice [1-3]: " missing_py_choice
 
   case "${missing_py_choice}" in
     1)
       echo "Please install python.org universal2 and re-run ./build.sh" >&2
       exit 2
       ;;
     2)
       if command -v python3 >/dev/null 2>&1; then
         PYTHON_UNIVERSAL2="$(command -v python3)"
       else
         echo "Error: python3 not found on PATH." >&2
         exit 2
       fi
       ;;
     3)
       read -r -p "Paste the full path to a working python3 executable: " python_override2
       if [[ -z "${python_override2}" ]]; then
         echo "Cancelled." >&2
         exit 2
       fi
       PYTHON_UNIVERSAL2="${python_override2}"
       ;;
     *)
       echo "Cancelled." >&2
       exit 2
       ;;
   esac
 
   if [[ ! -x "${PYTHON_UNIVERSAL2}" ]]; then
     echo "Error: Not executable: ${PYTHON_UNIVERSAL2}" >&2
     exit 2
   fi
 fi

 if [[ "${PYTHON_UNIVERSAL2}" != /Library/Frameworks/Python.framework/* ]]; then
   echo >&2
   echo "Warning: Selected Python does not look like a python.org Framework install:" >&2
   echo "  ${PYTHON_UNIVERSAL2}" >&2
   echo "Recommended universal2 installer: ${PYTHON_UNIVERSAL2_RELEASE_URL}" >&2
 fi

 if /usr/bin/lipo -archs "${PYTHON_UNIVERSAL2}" >/dev/null 2>&1; then
   py_archs=$(/usr/bin/lipo -archs "${PYTHON_UNIVERSAL2}" 2>/dev/null || true)
   if [[ "${py_archs}" != *"arm64"* || "${py_archs}" != *"x86_64"* ]]; then
     echo >&2
     echo "Warning: Selected Python does not appear to be universal2 (missing arm64/x86_64):" >&2
     echo "  ${PYTHON_UNIVERSAL2}" >&2
     echo "  archs: ${py_archs}" >&2
     echo "Recommended universal2 installer: ${PYTHON_UNIVERSAL2_RELEASE_URL}" >&2
   fi
 fi
 
 # If we are building Intel on Apple Silicon, ensure the selected Python can run under
 # Rosetta (x86_64). This is the main reason we recommend python.org universal2.
 if [[ ( "${TARGET}" == "intel" || "${TARGET}" == "intel_compat" ) && "$(uname -m)" == "arm64" ]]; then
   if ! arch -x86_64 "${PYTHON_UNIVERSAL2}" -c "import platform; print(platform.machine())" >/dev/null 2>&1; then
     echo >&2
     echo "Error: The selected Python cannot run as x86_64 under Rosetta:" >&2
     echo "  ${PYTHON_UNIVERSAL2}" >&2
     echo "For Intel builds on Apple Silicon, install python.org universal2 Python ${PYTHON_UNIVERSAL2_RECOMMENDED_VERSION} and re-run." >&2
     echo "Download: ${PYTHON_UNIVERSAL2_RELEASE_URL}" >&2
     echo "More downloads: https://www.python.org/downloads/macos/" >&2
     exit 2
   fi
 fi
 
 echo
 echo "Summary:"
 echo "  Target: ${TARGET}"
 echo "  Action: ${MODE}"
 echo "  Python: ${PYTHON_UNIVERSAL2}"
 read -r -p "Continue? [y/N]: " confirm
 if [[ "${confirm}" != "y" && "${confirm}" != "Y" ]]; then
   echo "Cancelled."
   exit 0
 fi
 
 HOST_ARCH="$(uname -m)"
 
 if [[ "${TARGET}" == "intel" ]]; then
   VENV_DIR="${VENV_INTEL}"
   REQUIREMENTS_FILE="requirements.txt"
   NUITKA_ARCH_FLAG="--macos-target-arch=x86_64"
   BUILD_DIR="build_intel"
 elif [[ "${TARGET}" == "intel_compat" ]]; then
   VENV_DIR="${VENV_INTEL_COMPAT}"
   REQUIREMENTS_FILE="${INTEL_COMPAT_REQUIREMENTS_FILE}"
   NUITKA_ARCH_FLAG="--macos-target-arch=x86_64"
   BUILD_DIR="build_intel_compat"
 else
   VENV_DIR="${VENV_APP}"
   REQUIREMENTS_FILE="requirements.txt"
   NUITKA_ARCH_FLAG="--macos-target-arch=arm64"
   BUILD_DIR="build_apple"
 fi
 
 need_rosetta=0
 if [[ ( "${TARGET}" == "intel" || "${TARGET}" == "intel_compat" ) && "${HOST_ARCH}" == "arm64" ]]; then
   need_rosetta=1
 fi

 confirm_delete_dir() {
   local dir="$1"
   if [[ -d "${dir}" ]] && [[ -n "$(/bin/ls -A "${dir}" 2>/dev/null)" ]]; then
     echo
     echo "Warning: '${dir}' already exists and is not empty."
     read -r -p "Delete it now? [y/N]: " del
     if [[ "${del}" != "y" && "${del}" != "Y" ]]; then
       echo "Cancelled."
       exit 2
     fi
   fi
 }
 
 run_in_target() {
   local cmd="$1"
   if [[ ${need_rosetta} -eq 1 ]]; then
     arch -x86_64 bash -c "${cmd}"
   else
     bash -c "${cmd}"
   fi
 }
 
 create_env() {
   echo "Creating venv: ${VENV_DIR}"
   if [[ -d "${VENV_DIR}" ]]; then
     echo "Deleting existing venv: ${VENV_DIR}"
   fi
   rm -rf "${VENV_DIR}"
   if [[ ${need_rosetta} -eq 1 ]]; then
     arch -x86_64 "${PYTHON_UNIVERSAL2}" -m venv "${VENV_DIR}"
   else
     "${PYTHON_UNIVERSAL2}" -m venv "${VENV_DIR}"
   fi
 
   if [[ "${TARGET}" == "apple" ]]; then
     run_in_target "source ${VENV_DIR}/bin/activate && pip install -U pip && pip install -r ${REQUIREMENTS_FILE} && pip install nuitka"
   else
     run_in_target "source ${VENV_DIR}/bin/activate && pip install -U pip && pip install --only-binary=:all: -r ${REQUIREMENTS_FILE} && pip install nuitka"
   fi
 }
 
 build_app() {
   echo "Building ${APP_NAME} v${VERSION} (target=${TARGET})"

   confirm_delete_dir "${BUILD_DIR}"
   rm -rf "${BUILD_DIR}"
   mkdir -p "${BUILD_DIR}"
 
   if [[ "${TARGET}" == "intel_compat" ]]; then
     run_in_target "export MACOSX_DEPLOYMENT_TARGET=${INTEL_DEPLOYMENT_TARGET}; source ${VENV_DIR}/bin/activate && python -m nuitka --version"
   else
     run_in_target "source ${VENV_DIR}/bin/activate && python -m nuitka --version"
   fi
 
   run_in_target "\
     source ${VENV_DIR}/bin/activate && \
     python -m nuitka \
       --standalone \
       --macos-create-app-bundle \
       ${NUITKA_ARCH_FLAG} \
       --macos-app-name=\"${APP_NAME}\" \
       --macos-app-version=\"${VERSION}\" \
       --macos-app-icon=explot.icns \
       --macos-signed-app-name=\"com.${APP_NAME}.app\" \
       --enable-plugin=tk-inter \
       --include-package=matplotlib.backends.backend_pdf \
       --output-filename=\"${APP_NAME}\" \
       --output-dir=\"${BUILD_DIR}\" \
       launch.py"
 
   if [[ ! -d "${BUILD_DIR}/launch.app" ]]; then
     echo "Error: Nuitka did not produce ${BUILD_DIR}/launch.app" >&2
     exit 1
   fi
 
   mv "${BUILD_DIR}/launch.app" "${BUILD_DIR}/${APP_NAME}.app"
 
   APP_BUNDLE="${BUILD_DIR}/${APP_NAME}.app"
   APP_EXE="${APP_BUNDLE}/Contents/MacOS/${APP_NAME}"
 
   if [[ -f "${APP_EXE}" ]]; then
     LIBPY_OLD=$(/usr/bin/otool -L "${APP_EXE}" | /usr/bin/awk '/libpython[0-9.]*\\.dylib/ {print $1; exit}')
 
     if [[ -n "${LIBPY_OLD}" ]]; then
       LIBPY_NAME=$(/usr/bin/basename "${LIBPY_OLD}")
       LIBPY_DEST="${APP_BUNDLE}/Contents/MacOS/${LIBPY_NAME}"
 
       if [[ ! -f "${LIBPY_DEST}" && -f "${LIBPY_OLD}" ]]; then
         /bin/cp "${LIBPY_OLD}" "${LIBPY_DEST}"
       fi
 
       if [[ -f "${LIBPY_DEST}" ]]; then
         for f in "${APP_BUNDLE}/Contents/MacOS/${APP_NAME}" "${APP_BUNDLE}/Contents/MacOS/"*.dylib "${APP_BUNDLE}/Contents/MacOS/"*.so; do
           if [[ -f "${f}" ]] && /usr/bin/otool -L "${f}" 2>/dev/null | /usr/bin/grep -q "${LIBPY_OLD}"; then
             /usr/bin/install_name_tool -change "${LIBPY_OLD}" "@executable_path/${LIBPY_NAME}" "${f}"
           fi
         done
       fi
     fi
   fi
 
   echo -e "Copying additional files..."
   cp -r pingouin "${BUILD_DIR}/${APP_NAME}.app/Contents/MacOS/"
   cp -r example_data.xlsx "${BUILD_DIR}/${APP_NAME}.app/Contents/MacOS/"
 
   echo -e "Build completed: ${BUILD_DIR}/${APP_NAME}.app"

   echo
   echo "Verify arch:"
   /usr/bin/file "${APP_EXE}" || true
   echo
   echo "Verify libpython linkage:"
   /usr/bin/otool -L "${APP_EXE}" | /usr/bin/grep -E "libpython[0-9.]*\\.dylib" || true

   echo
   echo "Summary:"
   echo "- Output: ${APP_BUNDLE}"
   echo "- Target: ${TARGET}"
   echo
   echo "Next: test-run the built app on this machine."
   if [[ ( "${TARGET}" == "intel" || "${TARGET}" == "intel_compat" ) && "$(uname -m)" == "arm64" ]]; then
     echo "Run (Rosetta): arch -x86_64 open \"${APP_BUNDLE}\""
   else
     echo "Run: open \"${APP_BUNDLE}\""
   fi

   echo
   read -r -p "Launch the app now? [y/N]: " run_now
   if [[ "${run_now}" == "y" || "${run_now}" == "Y" ]]; then
     if [[ ( "${TARGET}" == "intel" || "${TARGET}" == "intel_compat" ) && "$(uname -m)" == "arm64" ]]; then
       arch -x86_64 open "${APP_BUNDLE}"
     else
       open "${APP_BUNDLE}"
     fi
   fi
 }
 
 if [[ "${MODE}" == "env" || "${MODE}" == "all" ]]; then
   create_env
 fi
 
 if [[ "${MODE}" == "build" || "${MODE}" == "all" ]]; then
   if [[ ! -x "${VENV_DIR}/bin/python" ]]; then
     echo "Error: venv not found at ${VENV_DIR}. Please run the wizard again and choose 'Create/update environment'." >&2
     exit 1
   fi
   build_app
 fi

 exit 0