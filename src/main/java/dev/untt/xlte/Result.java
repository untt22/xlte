package dev.untt.xlte;

import java.util.function.Consumer;
import java.util.function.Function;

/**
 * Represents the result of an operation that can either succeed or fail.
 * This is a functional alternative to throwing exceptions for expected error cases.
 *
 * @param <T> The type of the success value
 */
public sealed interface Result<T> permits Result.Success, Result.Failure {

    /**
     * Represents a successful result containing a value.
     */
    record Success<T>(T value) implements Result<T> {}

    /**
     * Represents a failed result containing an error message.
     */
    record Failure<T>(String error) implements Result<T> {}

    /**
     * Transforms the success value using the given function.
     * If this is a Failure, returns a new Failure with the same error.
     *
     * @param mapper The function to apply to the success value
     * @param <U> The type of the mapped value
     * @return A new Result with the mapped value or the same error
     */
    default <U> Result<U> map(Function<T, U> mapper) {
        return switch (this) {
            case Success<T> s -> new Success<>(mapper.apply(s.value()));
            case Failure<T> f -> new Failure<>(f.error());
        };
    }

    /**
     * Chains another Result-producing operation.
     * If this is a Failure, returns a new Failure with the same error.
     *
     * @param mapper The function to apply to the success value
     * @param <U> The type of the new Result's value
     * @return The result of applying the mapper, or the same error
     */
    default <U> Result<U> flatMap(Function<T, Result<U>> mapper) {
        return switch (this) {
            case Success<T> s -> mapper.apply(s.value());
            case Failure<T> f -> new Failure<>(f.error());
        };
    }

    /**
     * Returns the value if this is a Success, otherwise returns the default value.
     *
     * @param defaultValue The value to return if this is a Failure
     * @return The success value or the default value
     */
    default T orElse(T defaultValue) {
        return switch (this) {
            case Success<T> s -> s.value();
            case Failure<T> f -> defaultValue;
        };
    }

    /**
     * Executes the consumer if this is a Failure.
     *
     * @param errorHandler The consumer to execute with the error message
     */
    default void ifFailure(Consumer<String> errorHandler) {
        if (this instanceof Failure<T> f) {
            errorHandler.accept(f.error());
        }
    }

    /**
     * Executes the consumer if this is a Success.
     *
     * @param successHandler The consumer to execute with the success value
     */
    default void ifSuccess(Consumer<T> successHandler) {
        if (this instanceof Success<T> s) {
            successHandler.accept(s.value());
        }
    }

    /**
     * Returns true if this is a Success.
     */
    default boolean isSuccess() {
        return this instanceof Success<T>;
    }

    /**
     * Returns true if this is a Failure.
     */
    default boolean isFailure() {
        return this instanceof Failure<T>;
    }
}
