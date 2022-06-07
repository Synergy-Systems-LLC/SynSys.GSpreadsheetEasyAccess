class AuthenticationTimedOutException(Exception):
    """Represents an exception thrown due to authentication timeout."""
    pass


class OAuthSheetsScopeException(Exception):
    """Represents an exception throwndue to a mismatch between the scopes
    of the application and the requests being sent."""
    pass


class UserCanceledAuthenticationException(Exception):
    """Represents an exception thrown because the user canceled the authentication process."""
    pass
