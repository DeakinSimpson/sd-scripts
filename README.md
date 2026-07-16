# SD Scripts

This is a Repo that can be used by SD to store scripts and work on them collaboratively

To run powershell scripts use the following command where <script location> is the location of the script:

    Get-Content -Raw .\<script location> | Invoke-Expression
