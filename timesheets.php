<?php

declare(strict_types=1);

use PhpdevUk\Timesheets\ReadFilter;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\Console\Application;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;

require_once __DIR__ . '/vendor/autoload.php';

const DATE_TIME_FORMAT = 'd/m/Y H:i:s A';

$application = new Application();

$application
    ->register('client')
    ->addArgument(
        'client',
        InputArgument::REQUIRED
    )
    ->addArgument(
        'files',
        InputArgument::IS_ARRAY | InputArgument::REQUIRED
    )
    ->addOption(
        'start-date',
        'start',
        InputOption::VALUE_REQUIRED,
        'Start Date (inclusive)',
        '2000-01-01'
    )
    ->addOption(
        'end-date',
        'end',
        InputOption::VALUE_REQUIRED,
        'End Date (inclusive)',
        '2100-01-01'
    )
    ->setCode(function (InputInterface $input, OutputInterface $output): int {
        $clientNameFilter = $input->getArgument('client');
        $files = $input->getArgument('files');
        $startDateFilter = $input->getOption('start-date');
        $endDateFilter = $input->getOption('end-date');

        $startDateTimeFilter = new DateTimeImmutable($startDateFilter . ' 00:00:00');
        $endDateTimeFilter = new DateTimeImmutable($endDateFilter . ' 00:00:00');
        // Add one day to end date to include the whole day
        $endDateTimeFilter = $endDateTimeFilter->modify('+1 day');

        $workChunks = [];

        foreach ($files as $file) {
            if (!file_exists($file)) {
                $output->writeln('File does not exist: ' . $file);
                return Command::FAILURE;
            }

            if (!is_file($file)) {
                $output->writeln('Not a file: ' . $file);
                return Command::FAILURE;
            }

            $readFilter = new ReadFilter('A', 'E');

            $reader = IOFactory::createReaderForFile($file);
            $reader->setReadDataOnly(false); // This must be false for formatted data to be returned
            $reader->setReadEmptyCells(false); // We only need non-empty cells, can guarantee each row will be populated
            $reader->setReadFilter($readFilter);
            $spreadsheet = $reader->load($file);
            $sheetCount = $spreadsheet->getSheetCount();

            for ($s = 0; $s < $sheetCount; $s++) {
                $sheet = $spreadsheet->getSheet($s);

                if ($output->isVerbose()) {
                    $output->writeln($sheet->getTitle());
                }

                $highestRow = $sheet->getHighestDataRow();

                for ($row = 2; $row <= $highestRow; $row++) {
                    $clientName = $sheet->getCell("E{$row}")->getFormattedValue();

                    if ($clientName == $clientNameFilter) {
                        $date = $sheet->getCell("A{$row}")->getFormattedValue();
                        $startTime = $sheet->getCell("B{$row}")->getFormattedValue();
                        $endTime = $sheet->getCell("C{$row}")->getFormattedValue();

                        $startDateTime = DateTimeImmutable::createFromFormat(DATE_TIME_FORMAT, "$date $startTime");
                        $endDateTime = DateTimeImmutable::createFromFormat(DATE_TIME_FORMAT, "$date $endTime");

                        if ($startDateTime >= $startDateTimeFilter && $startDateTime < $endDateTimeFilter) {
                            $diff = $startDateTime->diff($endDateTime);

                            $workChunks[] = [
                                'hours' => $diff->h,
                                'minutes' => $diff->i,
                            ];

                            if ($output->isVerbose()) {
                                $output->writeln("{$diff->h}h{$diff->i}m");
                            }
                        }
                    }
                }
            }
        }

        $totalHours = array_sum(array_column($workChunks, 'hours'));
        $totalMinutes = array_sum(array_column($workChunks, 'minutes'));

        // Convert minutes to hours
        while ($totalMinutes >= 60) {
            $totalHours++;
            $totalMinutes -= 60;
        }

        $output->writeln("Total hours: $totalHours");
        $output->writeln("Total minutes: $totalMinutes");

        return Command::SUCCESS;
    });

$application->run();
