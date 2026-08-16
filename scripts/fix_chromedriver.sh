#!/bin/bash
# Remove a specific incompatible ChromeDriver cache entry so Selenium can fetch a matching version.
rm -rf ~/.cache/selenium/chromedriver/linux64/146.0.7680.165 2>/dev/null || true
echo "ChromeDriver .165 removed if it existed"
