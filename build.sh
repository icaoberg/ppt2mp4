#!/usr/bin/env bash
set -euo pipefail

IMAGE="singularity-ppt2movie.sif"
DEF="ppt2movie.def"

if [[ ! -f "$DEF" ]]; then
    echo "Error: definition file '$DEF' not found."
    exit 1
fi

if [[ -f "$IMAGE" ]]; then
    echo "Removing existing image: $IMAGE"
    rm -f "$IMAGE"
fi

echo "Building Singularity image: $IMAGE"
sudo singularity build "$IMAGE" "$DEF"

echo "Done. Run with:"
echo "  singularity run $IMAGE <presentation.pptx> [options]"
