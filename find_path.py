def pathsplit(url):
    url = url
    url = url.split('/')
    length = len(url)
    index = length - 1
    filepath = url[index]
    return filepath

