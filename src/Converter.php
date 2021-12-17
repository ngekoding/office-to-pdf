<?php

namespace ngekoding\OfficeToPdf;

use Symfony\Component\Process\Exception\ProcessFailedException;
use Symfony\Component\Process\Process;

class Converter
{
    private $libreOfficePath;
    private $baseCommand = '%s --headless --convert-to pdf --outdir %s %s';

    public function __construct($libreOfficePath = NULL)
    {
        if (empty($libreOfficePath)) {
            $this->libreOfficePath = $this->getLibreOfficePath();
        } else {
            $this->libreOfficePath = $libreOfficePath;
        }
    }

    /**
     * Get libre office executable path based on current operating system
     * @return string Libre office executable path
     */
    public function getLibreOfficePath()
    {
        if (in_array(PHP_OS, ['Windows', 'WINNT', 'WIN32'])) {
            return 'soffice';
        } elseif (PHP_OS === 'Darwin') {
            return '/Applications/LibreOffice.app/Contents/MacOS/soffice';
        } else {
            return 'libreoffice';
        }
    }

    /**
     * Run converter process
     * @param $source       The file to convert
     * @param $destination  The converted file destination (directory or full path with filename)
     *                      Default filename will same with the source file,
     *                      You can change the filename by given a .pdf extension
     */
    public function convert($source, $destination = NULL)
    {
        $outdir = '.';
        $filename = NULL;

        // Configure output directory and filename
        if (!empty($destination)) {
            $pathinfo = pathinfo($destination);
            $extension = isset($pathinfo['extension']) ? $pathinfo['extension'] : NULL;

            if ($extension === 'pdf') {
                $outdir = $pathinfo['dirname'];
                $filename = $pathinfo['filename'].'.pdf';
            } else {
                $outdir = $destination;
            }
        }

        $command = sprintf($this->baseCommand, $this->libreOfficePath, $outdir, $source);
        
        $process = new Process($command);
        $process->run();

        if (!$process->isSuccessful()) {
            throw new ProcessFailedException($process);
        }

        // Renaming the output file if needed
        if (!empty($filename)) {
            $pathinfoSource = pathinfo($source);

            @rename(
                $outdir.'/'.$pathinfoSource['filename'].'.'.$pathinfoSource['extension'],
                $outdir.'/'.$filename
            );
        }
    }
}
