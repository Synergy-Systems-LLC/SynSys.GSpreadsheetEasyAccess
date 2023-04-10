from System import Byte


class OAuthSheetsScope(object):
    """OAuth scopes restrict the actions your application can perform on behalf of the end user.

    Set the minimum size required for your use case.\n
    The class implements only some of the areas from \
    [Google Sheets API v4](https://developers.google.com/identity/protocols/oauth2/scopes#sheets).
    """

    @property
    @classmethod
    def ViewAccess(cls):
        """Access to view Google spreadsheet."""
        return OAuthSheetsScope()

    @property
    @classmethod
    def FullAccess(cls):
        """Access to view, edit, create and delete Google spreadsheet."""
        return OAuthSheetsScope()


class SheetsService(object):
    pass


class Principal(object):
    """Principal is an object that can be granted access to a resource.

    See [Authentication Overview](https://cloud.google.com/docs/authentication) for details.
    """

    def GetSheetsService(self):
        """Return an object representing Google Sheets service.

        It [Google.Apis.Sheets.v4 library object](https://developers.google.com/resources/api-libraries/documentation/sheets/v4/csharp/latest/classGoogle_1_1Apis_1_1Sheets_1_1v4_1_1SheetsService.html)
        through which work with Google Sheets API takes place.
        """  # noqa
        return SheetsService()


class ServiceAccount(Principal):
    """Service accounts are managed by \
    [IAM](https://cloud.google.com/iam/docs/understanding-service-accounts),
    and represent non-human users.

    These are for scenarios where your application needs to access resources or
    perform actions on their own,
    such as running App Engine applications or interacting with Compute Engine instances.\n
    For more information, see \
    [Authenticate as a service account](https://cloud.google.com/docs/authentication/api-keys)
    """

    def __init__(self, apiKey):
        # type: (str) -> None
        """Authenticate the application as a service account."""
        pass


class UserAccount(Principal):
    """User accounts are managed as Google accounts, they represent a developer, administrator,
    or any other person who interacts with Google Cloud.

    These are for scenarios where your application needs to access resources as a human user.<br/>
    For more information, see \
    [Authenticate as an end user](https://cloud.google.com/docs/authentication/end-user).
    """

    def __init__(self, credentials, scope):
        # type: (list[Byte], OAuthSheetsScope) -> None
        """Authenticate as a human user."""
        pass

    @property
    def CancellationSeconds(self):
        """Time to try to connect to the application on the Google Cloud Platform in seconds.

        The default value is 30 seconds.
        """
        return int()

    @CancellationSeconds.setter
    def CancellationSeconds(self, value):
        # type: (int) -> None
        """Time to try to connect to the application on the Google Cloud Platform in seconds.

        The default value is 30 seconds.
        """
        pass
