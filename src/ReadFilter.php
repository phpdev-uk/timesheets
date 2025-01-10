<?php

declare(strict_types=1);

namespace PhpdevUk\Timesheets;

use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

class ReadFilter implements IReadFilter
{
    protected string $startColumn;
    protected string $endColumn;

    public function __construct(string $startColumn, string $endColumn)
    {
        $this->startColumn = $startColumn;
        $this->endColumn = $endColumn;
    }

    public function readCell(string $columnAddress, int $row, string $worksheetName = ''): bool
    {
        return in_array($columnAddress, range($this->startColumn, $this->endColumn));
    }
}