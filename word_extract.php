<?php

/**
 * Description: Extract all images from Microsoft Word file. A separate directory will be created
 *              for each DOCX file. BMP and TIFF formats will be converted
 *              to JPG format.
 *
 * Usage: php word_extract.php <filename.docx>
 * Requires Imagick for BMP and TIFF conversion to JPG
 *
 * Version: 1.0
 * Author: Daniel Caspi
 * Author URI: http://www.element26.net/
 * Copyright (C) 2015 Daniel Caspi and Element TwentySix
 *
 *    This program is free software: you can redistribute it and/or modify
 *    it under the terms of the GNU General Public License as published by
 *    the Free Software Foundation, either version 3 of the License, or
 *    (at your option) any later version.
 *
 *    This program is distributed in the hope that it will be useful,
 *    but WITHOUT ANY WARRANTY; without even the implied warranty of
 *    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *    GNU General Public License for more details.
 *
 *    You should have received a copy of the GNU General Public License
 *    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *
 **/

// Check for valid input
if (!isset($argv[1]))
    die("No file name provided!\n");
if (!is_file($argv[1]))
    die("Invalid filename $argv[1] provided!\n");

// Name of the document file
$document = $argv[1];

readZippedImages($document);

// Function to extract images
function readZippedImages($filename)
{
    // Set output directory
    $document_dir = getcwd() . '/' . preg_replace('/\.[^.]+$/', '', $filename) . '/';

    // Create a new ZIP archive object
    $zip = new ZipArchive;

    // Open the received archive file
    if (true === $zip->open($filename)) {

         // Create directory with same name
        if (!is_dir($document_dir)) mkdir($document_dir, 0777, true);

        for ($i = 0; $i < $zip->numFiles; $i++) {

            // Loop via all the files to check for image files
            $zip_element = $zip->statIndex($i);
            $image       = $zip->getFromIndex($i);

            // Retrieve images only
            if (preg_match("@^word/media/.*\.(jpg|jpeg|png|gif|emf)$@i", $zip_element['name'])) {

                $filename = preg_replace('/^word\/media\//i', '', $zip_element['name']);
                $filename = preg_replace('/\.jpeg$/i', '.jpg', $filename);
                file_put_contents($document_dir . $filename, $image, LOCK_EX);
                echo "Extracted " . $zip_element['name'] . "\n";

            } elseif (preg_match("@^word/media/.*\.(bmp|tif|tiff)$@i", $zip_element['name'])) {

                // Compress BMP and TIFF to JPG
                // If compression is undesired, simply add bmp|tif|tiff to the preg_match function on line 59

                $filename = preg_replace('/^word\/media\//i', '', $zip_element['name']);
                $filename = preg_replace('/\.[^.]+$/', '.jpg', $filename);

                $imagick = new Imagick();
                $imagick->readImageBlob($image);
                $imagick->trimImage(0);

                $imagick->setCompression(Imagick::COMPRESSION_JPEG);
                $imagick->setCompressionQuality(100);
                $imagick->setImageFormat('jpg');

                $imagick->writeImage($document_dir . $filename);
                $imagick->clear();
                $imagick->destroy();

                echo "Converted " . $zip_element['name'] . " to JPG\n";

            } elseif (preg_match("@^word/media/@i", $zip_element['name'])) {

                $filename = preg_replace('/^word\/media\//i', '', $zip_element['name']);
                file_put_contents($document_dir . $filename, $image, LOCK_EX);
                echo "Unknown format " . $zip_element['name'] . "\n";

            }
        }
        $zip->close();
    }
}

?>
