#!/usr/bin/env bash
set -euo pipefail

IMAGE="ppt2vid.sif"
DEF="ppt2vid.def"

if [[ ! -f "$DEF" ]]; then
    echo "Error: definition file '$DEF' not found."
    exit 1
fi

if [[ -f "$IMAGE" ]]; then
    echo "Removing existing image: $IMAGE"
    rm -f "$IMAGE"
fi

echo "Building Singularity image: $IMAGE"
singularity build "$IMAGE" "$DEF"

echo "Done. Run with:"
echo "  singularity run $IMAGE <presentation.pptx> [options]"
