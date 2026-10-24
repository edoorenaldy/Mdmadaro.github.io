# Material Catalogue Images

The material-image build maps each Material Number to `assets/images/<Material Number>.jpg`.

- The build downloads representative product photos from web image sources.
- Images are converted to valid RGB JPEG files and compressed to a maximum of 3,000 bytes each.
- `assets/images/MATERIAL_IMAGE_SOURCES.csv` records the material number, description, source image URL, file size, and build status.
- Photos are representative matches by description/category unless a product/model match is independently confirmed. Do not treat a representative image as proof of the exact supplier model.

The automated build is defined in `.github/workflows/material-images.yml` and `scripts/build_material_images.py`.
