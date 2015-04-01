<?php
namespace mbright\ExcelOutput;

/**
 * Class Parambag
 *
 * Simple \ArrayAccess class to shove key value pairs into.
 *
 * @package mbright\ExcelOutput
 */
class Parambag implements \ArrayAccess
{
    /**
     * @var array
     */
    protected $params = array();

    /**
     * Create a new parameter bag instance.
     *
     * @param array $params
     */
    public function __construct(array $params = array())
    {
        $this->params = $params;
    }

    /**
     * Get parameter values out of the bag.
     *
     * @param string $key
     * @return mixed
     */
    public function get($key)
    {
        return $this->__get($key);
    }

    /**
     * Get parameter values out of the bag via the property access magic method.
     *
     * @param string $key
     * @return mixed
     */
    public function __get($key)
    {
        return isset($this->params[$key]) ? $this->params[$key] : null;
    }

    /**
     * Check if a param exists in the bag via an isset() check on the property.
     *
     * @param string $key
     * @return bool
     */
    public function __isset($key)
    {
        return isset($this->params[$key]);
    }

    /**
     * Check if a param exists in the bag via an isset() and array access.
     *
     * @param string $key
     * @return bool
     */
    public function offsetExists($key)
    {
        return $this->__isset($key);
    }

    /**
     * Get parameter values out of the bag via array access.
     *
     * @param string $key
     * @return mixed
     */
    public function offsetGet($key)
    {
        return $this->__get($key);
    }

    /**
     * Sets a particular key and value.
     *
     * @param string $key
     * @param mixed  $value
     * @throws \LogicException
     * @return void
     */
    public function offsetSet($key, $value)
    {
        $this->params[$key] = $value;
    }

    /**
     * Unset() a particular key.
     *
     * @param string $key
     * @throws \LogicException
     * @return void
     */
    public function offsetUnset($key)
    {
        unset($this->params[$key]);
    }
}