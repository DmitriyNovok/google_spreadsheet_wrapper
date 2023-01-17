<?php

namespace src\GoogleSpreadsheet;

use Google\Exception;
use Google_Service_Sheets;

abstract class GoogleSpreadsheetWrapper
{
    /**
     * Scopes to be requested.
     *
     * @var string
     */
    protected string $scope;

    /**
     * Auth config configuration json
     */
    protected $credentials;

    /**
     * @deprecated
     * Access token used for requests.
     *
     * @var string the configuration json
     */
    protected string $tokenPath;

    /**
     * Range cells in list.
     *
     * @var string Example !A1:E1
     */
    protected string $range;

    /**
     * @var object
     */
    protected object $service;

    /**
     * @var mixed
     */
    protected mixed $client;

    /**
     * ID spreadsheet.
     *
     * @var string
     */
    protected string $spreadsheetId;

    /**
     * ID sheet.
     *
     * @var string
     */
    protected string $sheetId;

    /**
     * Sheet title spreadsheet list.
     *
     * @var string
     */
    protected string $sheetTitle;

    /**
     * Set application name.
     *
     * @var string
     */
    protected string $applicationName;

    /**
     *
     * @throws \Exception
     */
    public function __construct()
    {
        return $this->getClient();
    }

    /**
     *
     * @throws \Exception
     */
    private function getClient()
    {
        if (empty($this->credentials)) {
            throw new \Exception('Required credentials config.');
        }

        $client = new \Google_Client();
        $client->setApplicationName($this->applicationName);
        $client->setScopes($this->scope);
        $client->setAuthConfig($this->credentials);
        $client->setAccessType('offline');

        $this->setClient($client);

        return $this->client;
    }

    /**
     * @return mixed
     * @throws Exception
     *
     * @deprecated
     */
    private function getApiClient(): mixed
    {
        if (empty($this->credentials)) {
            throw new \Exception('File does not exist.');
        }

        if (is_string($this->tokenPath) && empty($this->tokenPath)) {
            throw new \Exception('File does not exist.');
        }

        if (empty($this->applicationName)) {
            throw new \Exception('Application name was not set.');
        }

        if (empty($this->scope)) {
            throw new \Exception('Scopes was not set.');
        }

        $googleClient = new \Google_Client();
        $googleClient->setApplicationName($this->applicationName);
        $googleClient->setScopes($this->scope);
        $googleClient->setAuthConfig($this->credentials);
        $googleClient->setAccessType('offline');
        $googleClient->setPrompt('select_account consent');

        // Load previously authorized token from a file, if it exists.
        // The file token.json stores the user's access and refresh tokens, and is
        // created automatically when the authorization flow completes for the first
        // time.
        $tokenPath = $this->tokenPath;
        if (file_exists($tokenPath)) {
            $accessToken = json_decode(file_get_contents($tokenPath), true);
            $googleClient->setAccessToken($accessToken);
        }

        // If there is no previous token or it's expired.
        if ($googleClient->isAccessTokenExpired()) {
            // Refresh the token if possible, else fetch a new one.
            if ($googleClient->getRefreshToken()) {
                $googleClient->fetchAccessTokenWithRefreshToken($googleClient->getRefreshToken());
            } else {
                // Request authorization from the user.
                $authUrl = $googleClient->createAuthUrl();
                printf("Open the following link in your browser:\n%s\n", $authUrl);
                echo 'Enter verification code: ';
                $authCode = trim(fgets(\STDIN));

                // Exchange authorization code for an access token.
                $accessToken = $googleClient->fetchAccessTokenWithAuthCode($authCode);
                $googleClient->setAccessToken($accessToken);

                // Check to see if there was an error.
                if (array_key_exists('error', $accessToken)) {
                    throw new \Exception(implode(', ', $accessToken));
                }
            }

            // Save the token to a file.
            if (!file_exists(dirname($tokenPath))) {
                mkdir(dirname($tokenPath), 0700, true);
            }

            file_put_contents($tokenPath, json_encode($googleClient->getAccessToken()));
        }

        $this->setClient($googleClient);

        return $this->client;
    }

    /**
     * Get API client and the service object.
     *
     * @return Google_Service_Sheets
     */
    private function getServiceObject(): Google_Service_Sheets
    {
        $this->service = new Google_Service_Sheets($this->client);

        return $this->service;
    }

    /**
     * Set required the necessary credentials your spreadsheet.
     *
     * @return mixed
     */
    abstract protected function setUp(): mixed;

    /**
     * Add empty row to up list.
     */
    public function insertRequestRange(): void
    {
        $dimensionRange = new \Google_Service_Sheets_DimensionRange();
        $dimensionRange->setSheetId($this->sheetId);
        $dimensionRange->setDimension('ROWS');
        $dimensionRange->setStartIndex(1);
        $dimensionRange->setEndIndex(2);

        $dimensionRequest = new \Google_Service_Sheets_InsertDimensionRequest();
        $dimensionRequest->setRange($dimensionRange);
        $dimensionRequest->setInheritFromBefore(false);

        $sheetsRequest = new \Google_Service_Sheets_Request();
        $sheetsRequest->setInsertDimension($dimensionRequest);

        $updateSpreadsheetRequest = new \Google_Service_Sheets_BatchUpdateSpreadsheetRequest();
        $updateSpreadsheetRequest->setRequests([$sheetsRequest]);

        $this->getServiceObject()
            ->spreadsheets
            ->batchUpdate(
                $this->spreadsheetId,
                $updateSpreadsheetRequest,
                []
            );
    }

    /**
     * Insert new row with values.
     *
     * @param array $data
     */
    public function insertRow(array $data): void
    {
        $values = [$data];

        if (!empty($values)) {
            $this->getServiceObject()
                ->spreadsheets_values
                ->append(
                    $this->spreadsheetId,
                    $this->sheetTitle.$this->range,
                    new \Google_Service_Sheets_ValueRange(['values' => $values]),
                    ['valueInputOption' => 'RAW']
                );

            /**
             * After insert new row add empty row to up list for new rows
             * Can hide if not need
             */
            $this->insertRequestRange();
        }
    }

    /**
     * Update row.
     *
     * @param string $spreadsheetId
     * @param string $sheetTitle
     * @param string $rangeCells !A1:E1
     * @param array $values
     *
     * @throws \Exception
     */
    public function updateRow(string $spreadsheetId, string $sheetTitle, string $rangeCells, array $values): void
    {
        if (empty($sheetTitle)) {
            throw new \Exception('Sheet title is not must be empty.');
        }

        $sheetTitle = is_string($sheetTitle) ? $sheetTitle : (string) $sheetTitle;
        $rangeCells = is_string($rangeCells) ? $rangeCells : (string) $rangeCells;

        $this->getServiceObject()
            ->spreadsheets_values
            ->update(
                $spreadsheetId ?: $this->spreadsheetId,
                $sheetTitle.$rangeCells,
                new \Google_Service_Sheets_ValueRange([
                    'values' => is_array($values) ? $values : [$values]
                ]),
                ['valueInputOption' => 'RAW']
            );
    }

    /**
     * Create new tab.
     *
     * @param string $spreadsheetId
     * @param string $title
     */
    public function createNewSheet(string $spreadsheetId, string $title): void
    {
        $batchUpdateSpreadsheetRequest = new \Google_Service_Sheets_BatchUpdateSpreadsheetRequest([
            'requests' => [
                'addSheet' => [
                    'properties' => [
                        'title' => $title
                    ]
                ]
            ]
        ]);

        $this->getServiceObject()
            ->spreadsheets
            ->batchUpdate(
                $spreadsheetId ?: $this->spreadsheetId,
                $batchUpdateSpreadsheetRequest
            );
    }

    /**
     * Get row from sheet in set range.
     *
     * @param string $spreadsheetId
     * @param string $sheetTitle
     * @param string $rangeCells !A1:E1
     *
     * @throws \Exception
     *
     * @return array
     */
    public function getSpreadsheetValues(string $spreadsheetId, string $sheetTitle, string $rangeCells): array
    {
        if (empty($sheetTitle)) {
            throw new \Exception('Sheet title is not must be empty.');
        }

        $sheetTitle = is_string($sheetTitle) ? $sheetTitle : (string) $sheetTitle;
        $rangeCells = is_string($rangeCells) ? $rangeCells : (string) $rangeCells;

        $response = $this->getServiceObject()
            ->spreadsheets_values
            ->get(
                $spreadsheetId ?: $this->spreadsheetId,
                $sheetTitle.$rangeCells
            );

        return $response->getValues();
    }

    /**
     * Remove row in sheet in set range.
     *
     * @param string $spreadsheetId
     * @param string $sheetTitle
     * @param string $rangeRemoveCells !A2:E2
     */
    public function removeRow(string $spreadsheetId, string $sheetTitle, string $rangeRemoveCells): void
    {
        $sheetTitle = is_string($sheetTitle) ? $sheetTitle : (string) $sheetTitle;
        $rangeRemoveCells = is_string($rangeRemoveCells) ? $rangeRemoveCells : (string) $rangeRemoveCells;

        $this->getServiceObject()
            ->spreadsheets_values
            ->clear(
                $spreadsheetId ?: $this->spreadsheetId,
                $sheetTitle.$rangeRemoveCells,
                new \Google_Service_Sheets_ClearValuesRequest()
            );
    }

    public function setClient($client): void
    {
        $this->client = $client;
    }

    public function setSpreadsheetId(string $spreadsheetId): void
    {
        $this->spreadsheetId = $spreadsheetId;
    }

    public function getClientApi()
    {
        return $this->client;
    }

    public function getSpreadsheetId(): string
    {
        return $this->spreadsheetId;
    }
}
