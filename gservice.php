<?php

namespace App\Services;

class GoogleDrive
{
    public function __construct()
    {

    }

    public function getClient($code=null)
    {
        $client = new Google_Client();
        $client->setAuthConfig(dirname(__FILE__).'/oauth-credential.json');
        $client->setScopes(Google_Service_Sheets::SPREADSHEETS);
        $client->setScopes(Google_Service_Drive::DRIVE);
        $client->setAccessType('offline');

        // Load previously authorized token from a file, if it exists.
        // The file token.json stores the user's access and refresh tokens, and is
        // created automatically when the authorization flow completes for the first
        // time.
        $tokenPath = dirname(__FILE__).'/token.json';
        if (file_exists($tokenPath)) {
            $accessToken = json_decode(file_get_contents($tokenPath), true);
            $client->setAccessToken($accessToken);
        }

        // If there is no previous token or it's expired.
        if ($client->isAccessTokenExpired()) {
            // Refresh the token if possible, else fetch a new one.
            if ($client->getRefreshToken()) {
                $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            } else {
                // Request authorization from the user.
                $authUrl = $client->createAuthUrl();
                header("Location:".filter_var($authUrl, FILTER_SANITIZE_URL));
            }

            if($code)
            {
                // Exchange authorization code for an access token.
                $accessToken = $client->fetchAccessTokenWithAuthCode($code);
                $client->setAccessToken($accessToken);

                // Check to see if there was an error.
                if (array_key_exists('error', $accessToken)) {
                    throw new Exception(join(', ', $accessToken));
                }

                // Save the token to a file.
                if (!file_exists(dirname($tokenPath))) {
                    mkdir(dirname($tokenPath), 0700, true);
                }
                file_put_contents($tokenPath, json_encode($client->getAccessToken()));
            }

        }
        return $client;
    }

    public function findFolderByName($folder_name, $client=null)
    {
        if(!$client)
        {
            $client = $this->getClient();
        }

        $service = new Google_Service_Drive($client);

        $data = array();
        try {
            $result = $service->files->listFiles(array(
                "q" => array(
                    "mimeType='application/vnd.google-apps.folder'",
                ),
                'spaces' => 'drive',
                'fields' => 'files(id, name)',
            ));

            foreach ($result->getFiles() as $folder) {
                if($folder->getName()==$folder_name)
                {
                    $data['folder_id']=$folder->getId();
                    $data['folder_name']=$folder->getName();
                    break;
                }
            }
        } catch (Exception $e) {}

        return $data;
    }

    public function createFile($args, $folder_id = null, $client=null){
        if(!$client)
        {
            $client = $this->client();
        }
        $service = new Google_Service_Drive($client);

		$name = $args['name'];
		$content = $args['content'];
		$mimeType = $args['mimeType'];
		$description = $args['description'];
        $folder_id = $folder_id ? $folder_id : 'root';
        $fileMetadata = new Google_Service_Drive_DriveFile([
            'name' 		  => $name,
			'description' => $description,
            'parents' 	  => array($folder_id)
        ]);

        $file = $service->files->create($fileMetadata, [
            'data' => $content,
            'mimeType' => $mimeType,
            'uploadType' => 'multipart',
            'fields' => 'id'
        ]);

        return $file->id;
    }

    function createFolder($folder_name, $client = null)
    {
        if(!$client)
        {
            $client = $this->getClient();
        }

        $service = new Google_Service_Drive($client);

        $fileMetadata = new Google_Service_Drive_DriveFile();
        $fileMetadata->setName($folder_name);
        $fileMetadata->setMimeType("application/vnd.google-apps.folder");
        $file = $service->files->create($fileMetadata, array(
            'fields' => 'id'
        ));
        return $file->getId();
    }

    public function setToken($code)
    {
        $client = $this->getClient($code);
        $redirect_uri = base_url('/report/index');
        header('Location: ' . filter_var($redirect_uri, FILTER_SANITIZE_URL));
    }

    public function uploadToClient($filesToSend, $folder_name, $file_name) {
        $client = $this->getClient();
        $service = new Google_Service_Drive($client);

        $folder_exists = $this->findFolderByName($folder_name);
        if (count($folder_exists)==0) {
            $folderId = $this->createFolder($folder_name);
        } else {
            $folderId = $folder_exists['folder_id'];
        }

        $fileMetadata = new Google_Service_Drive_DriveFile();
        $fileMetadata->setName($file_name);
        $fileMetadata->setParents(array($folderId));
        $fileMetadata->setMimeType('application/vnd.ms-excel');

        // $content = file_get_contents($filesToSend);
        $content = $this->getXLSData($filesToSend);
        $file = $service->files->create($fileMetadata, array(
                'data'       => $content,
                'uploadType' => 'multipart',
                'fields'     => 'id'
        ));

        return $file->id;
    }

    public function duplicateFile($file_source_id, $old_folder_id, $new_folder_id, $new_file_name, $client=null)
    {
        if(!$client)
        {
            $client = $this->getClient();
        }

        $service = new Google_Service_Drive($client);

        $file = new Google_Service_Drive_DriveFile([
            'name' => $new_file_name,
        ]);

        $cloned = $service->files->copy($file_source_id, $file);
        $fileId = $cloned->id;

        $params = array(
            'addParents'    => $new_folder_id,
            'removeParents' => $old_folder_id,
            'fields'        => 'id, parents',
        );

        $empty_meta_file = new Google_Service_Drive_DriveFile();
        $updated = $service->files->update($fileId, $empty_meta_file, $params);
        return $updated;
    }

    public function appendRow($file_source_id, $values, $client = null)
    {
        if(!$client)
        {
            $client = $this->getClient();
        }

        $service = new Google_Service_Sheets($client);

        $range = 'Sheet1';
        $val = [$values];

        $body = new Google_Service_Sheets_ValueRange([
            'values' => $val
        ]);

        $params = [
            'valueInputOption' => 'RAW'
        ];
        $insert = [
            "insertDataOption" => "INSERT_ROWS"
        ];
        $result = $service->spreadsheets_values->append(
            $file_source_id,
            $range,
            $body,
            $params,
            $insert
        );

        return $result;
    }

    public function clearValues($file_source_id, $client = null)
    {
        if(!$client)
        {
            $client = $this->getClient();
        }

        $service = new Google_Service_Sheets($client);

        $range = 'Sheet1!A3:BQ1000';
        $requestBody = new Google_Service_Sheets_ClearValuesRequest();
        return $service->spreadsheets_values->clear($file_source_id, $range, $requestBody);
    }

}
